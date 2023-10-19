module Purifications

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open HelperFunctions

// Function used to fill out purification Batch Records
let purificationStart (inputParams : string list) (purificationsForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    // Gets logged user for use in path and error logging
    let user = Environment.UserName
    //Console.WriteLine "Which reagents will you use?"
    let reagentsInput = "purifications" //Console.ReadLine ()
    //User Interface
    Console.WriteLine "What is the regen of the beads being used?"
    let regenNumber = Console.ReadLine ()
    Console.WriteLine "At which bench are you working?"
    let benchId = Console.ReadLine ()

    // Cycles through CS build IDs and fills out all reagent info from LIMS
    for param in inputParams do
              
            // Takes value of each cell of the row in which the input lies and stores it for use in filling out Word Doc
            let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)

            // Reading reagents from LIMS and loking for reagent box input
            let sspeLot = purificationReagentsList  reagentsInput myTools  2
            let sspeExp = purificationReagentsList  reagentsInput myTools  3
            let sspeUbd = purificationReagentsList  reagentsInput myTools  4
            let tweenLot = purificationReagentsList  reagentsInput myTools  5
            let tweenExp = purificationReagentsList  reagentsInput myTools  6
            let tweenUbd = purificationReagentsList  reagentsInput myTools  7
            let depcLot = purificationReagentsList  reagentsInput myTools  8
            let depcExp = purificationReagentsList  reagentsInput myTools  9
            let depcUbd = purificationReagentsList  reagentsInput myTools  10
            let oneXLot = purificationReagentsList  reagentsInput myTools  11
            let oneXExp = purificationReagentsList  reagentsInput myTools  12
            let oneXUbd = purificationReagentsList  reagentsInput myTools  13
            let halfXLot = purificationReagentsList  reagentsInput myTools  14
            let halfXExp = purificationReagentsList  reagentsInput myTools  15
            let halfXUbd = purificationReagentsList  reagentsInput myTools  16
            let hybId = purificationReagentsList  reagentsInput myTools  17
            let HybCalibrationDate = purificationReagentsList  reagentsInput myTools  18
            let nanoDropId = purificationReagentsList  reagentsInput myTools  19
            let nanoDropCalibration = purificationReagentsList  reagentsInput myTools  20

            // Reading bench equipment IDs from LIMS to find pipettes associated with bench specified by user
            let pipetteCal = purificationReagentsList benchId myTools  2
            let p1000 = purificationReagentsList benchId myTools  3
            let p1000Id = purificationReagentsList benchId myTools  4
            let p200 = purificationReagentsList benchId myTools  5
            let p200Id = purificationReagentsList benchId myTools  6
            let p20 = purificationReagentsList benchId myTools  7
            let p20Id = purificationReagentsList benchId myTools  8
            let p2000 = purificationReagentsList benchId myTools  9
            let p2000Id = purificationReagentsList benchId myTools  10
            let p2 = purificationReagentsList benchId myTools  11
            let p2Id = purificationReagentsList benchId myTools  12

            // Reads in copy of Batch Record template and starts processing
            let docArray = File.ReadAllBytes(purificationsForm)
            use docCopy = new MemoryStream(docArray)
            use purificationDocument = WordprocessingDocument.Open (docCopy, true)
            let purificationBody = purificationDocument.MainDocumentPart.Document.Body

            // Fills out CS identifying info
            (purificationCsInfoHeader purificationBody 1 5).Text <- lot + " " + csName
            (purificationCsInfoHeader purificationBody 1 13).Text <- geneNumber |> string
            (purificationCsInfoHeader purificationBody 1 19).Text <- scale |> string

            // Fills reagent and equipment info 
            (fillingPurificationLots purificationBody 0 1 2 0 0).Text <- lot
            (fillingPurificationLots purificationBody 0 1 3 0 0).Text <- "N/A"
            (fillingPurificationLots purificationBody 0 2 2 0 0).Text <- sspeLot
            (fillingPurificationLots purificationBody 0 2 3 1 0).Text <- sspeExp
            (fillingPurificationLots purificationBody 0 2 3 2 3).Text <- sspeUbd
            (fillingPurificationLots purificationBody 0 3 2 0 0).Text <- tweenLot
            (fillingPurificationLots purificationBody 0 3 3 1 0).Text <- tweenExp
            (fillingPurificationLots purificationBody 0 3 3 2 3).Text <- tweenUbd
            (fillingPurificationLots purificationBody 0 4 2 0 0).Text <- depcLot
            (fillingPurificationLots purificationBody 0 4 3 1 0).Text <- depcExp
            (fillingPurificationLots purificationBody 0 4 3 2 3).Text <- depcUbd
            (fillingPurificationLots purificationBody 0 5 0 1 2).Text <- regenNumber
            (fillingPurificationLots purificationBody 0 5 2 0 0).Text <- regenNumber
            (fillingPurificationLots purificationBody 0 5 3 0 0).Text <- "N/A"
            (fillingPurificationLots purificationBody 0 6 2 0 0).Text <- oneXLot
            (fillingPurificationLots purificationBody 0 6 3 1 0).Text <- oneXExp
            (fillingPurificationLots purificationBody 0 6 3 2 3).Text <- oneXUbd
            (fillingPurificationLots purificationBody 0 7 2 0 0).Text <- halfXLot
            (fillingPurificationLots purificationBody 0 7 3 1 0).Text <- halfXExp
            (fillingPurificationLots purificationBody 0 7 3 2 3).Text <- halfXUbd
            (fillingPurificationLots purificationBody 0 9 1 0 0).Text <- hybId
            (fillingPurificationLots purificationBody 0 9 2 0 0).Text <- HybCalibrationDate
            (fillingPurificationLots purificationBody 0 10 1 0 0).Text <- nanoDropId
            (fillingPurificationLots purificationBody 0 10 2 0 0).Text <- nanoDropCalibration
            (fillingPurificationLots purificationBody 0 12 0 0 0).Text <- p1000
            (fillingPurificationLots purificationBody 0 12 1 0 1).Text <- p1000Id
            (fillingPurificationLots purificationBody 0 12 2 0 0).Text <- pipetteCal
            (fillingPurificationLots purificationBody 0 13 0 0 0).Text <- p200
            (fillingPurificationLots purificationBody 0 13 1 0 1).Text <- p200Id
            (fillingPurificationLots purificationBody 0 13 2 0 0).Text <- pipetteCal
            (fillingPurificationLots purificationBody 0 14 0 0 0).Text <- p20
            (fillingPurificationLots purificationBody 0 14 1 0 1).Text <- p20Id
            (fillingPurificationLots purificationBody 0 14 2 0 0).Text <- pipetteCal
            (fillingPurificationLots purificationBody 0 15 0 0 0).Text <- p2000
            (fillingPurificationLots purificationBody 0 15 1 0 1).Text <- p2000Id
            (fillingPurificationLots purificationBody 0 15 2 0 0).Text <- pipetteCal
            (fillingPurificationLots purificationBody 0 16 0 0 0).Text <- p2
            (fillingPurificationLots purificationBody 0 16 1 0 1).Text <- p2Id
            (fillingPurificationLots purificationBody 0 16 2 0 0).Text <- pipetteCal


            // Calculates reagent volumes needed for chemistry reactions
            let volume = System.Math.Ceiling(theoreticalVolume (scale |> float) (geneNumber |> float))
            let oneXVolume = System.Math.Round(((volume * 0.95) * 1.2), 1) 
            let sspeVolume = System.Math.Round((oneXVolume / 20.0), 1)
            let tweenVolume = System.Math.Round((oneXVolume / 100.0), 1)
            let water = System.Math.Round((oneXVolume - volume - sspeVolume - tweenVolume), 1)
            let calculatingBeadVolume = (oneXVolume * 2.0)
            let beads = 
                if(calculatingBeadVolume % 100.0) = 0.0 then
                    (oneXVolume * 2.0).ToString()
                else
                    ((100.0 - (calculatingBeadVolume % 100.0)) + calculatingBeadVolume).ToString()

            let buffer = 
                if (geneNumber |> float) < 100.0 then 
                    (((((scale |> float) - 1.0) * 1000.0) * 0.6) / (618.0 / (geneNumber |> float)))
                      
                elif (geneNumber |> float) < 400.0 then 
                    ((((((scale |> float) - 1.0) * 1000.0) * 0.6)) / 5.0)
                else
                    ((((((scale |> float) - 1.0) * 1000.0) * 0.6)) / 3.6)
            let elutionBuffer = (System.Math.Ceiling(buffer)).ToString()

            // Fills out Batch Record with reagent quantities for purification chemistry
            (writingCalculations purificationBody 1 4 1 8 2).Text <- volume.ToString()
            (writingCalculations purificationBody 1 5 1 0 1).Text <- oneXVolume.ToString()
            (writingCalculations purificationBody 1 8 1 0 5).Text <- volume.ToString()
            (writingCalculations purificationBody 1 8 1 1 6).Text <- sspeVolume.ToString()
            (writingCalculations purificationBody 1 8 1 2 5).Text <- tweenVolume.ToString()
            (writingCalculations purificationBody 1 8 1 3 2).Text <- water.ToString()
            (writingCalculations purificationBody 1 8 1 5 3).Text <- oneXVolume.ToString()
            (writingCalculations purificationBody 1 9 1 0 2).Text <- elutionBuffer
            (writingCalculations purificationBody 1 10 1 1 2).Text <- beads
            (writingCalculations purificationBody 1 13 1 0 2).Text <-beads
            (writingCalculations purificationBody 1 19 1 2 2).Text <- beads
            (writingCalculations purificationBody 1 22 1 2 2).Text <- beads
            (writingCalculations purificationBody 1 25 1 2 2).Text <- elutionBuffer

            // Saves document in temp directory for printing and subsequent deletion
            let purificationFormPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Purification Form" + ".docx"
            purificationDocument.SaveAs(purificationFormPath).Close() |> ignore
