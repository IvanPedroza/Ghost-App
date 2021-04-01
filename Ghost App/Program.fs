// Learn more about F# at http://fsharp.org

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open System.Diagnostics
open Sentry
open Sentry.Integrations
open HelperFunctions

let pathsList (user : string) (param : string) = 
    [
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Request Form" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Ligation Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Gel Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Zag Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Purification Form" + ".docx"
    ]


let printDocuments (path : string) =
    let printing = new Process()
    printing.StartInfo.FileName <- path
    printing.StartInfo.Verb <- "Print"
    printing.StartInfo.CreateNoWindow <- true
    printing.StartInfo.UseShellExecute <- true
    printing.EnableRaisingEvents <- true
    printing.Start() |> ignore
    printing.WaitForExit(10000)

    File.Delete(path)

[<EntryPoint>]
let main argv =
    use __ = SentrySdk.Init (fun o ->
        o.Dsn <-  "https://d1553cf78c164e5d9813ca11cc417d80@o561151.ingest.sentry.io/5697684"
        o.SendDefaultPii <- true
        o.StackTraceMode
        o.AttachStacktrace <- true
        o.ShutdownTimeout <- TimeSpan.FromSeconds 10.0 
        o.MaxBreadcrumbs <- 50 
        )
    //Reading in exel doc
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial  
    
    //User interface - input will take string as input
    Console.WriteLine "What GP will you work on?"
    let input = Console.ReadLine ()
    let inputSplit = input.Split(' ')
    let inputParams = [for i in inputSplit do i.ToUpper()]

    Console.WriteLine "Which process are you conducting?"
    let processInput = Console.ReadLine ()
    
    //Reading in excel file from path on Reporter Probe sheet
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let fileInfo = new FileInfo("W:/Production/Reporter requests/CODESET TRACKERS/CodeSet Status.xlsx")
    use package = new ExcelPackage (fileInfo)
    let ghost = package.Workbook.Worksheets.["Upstream - GP"]

    //Reading in reagents excel book
    let reagentsInfo = new FileInfo("C:/Users/ipedroza/Desktop/reagentsandtools.xlsx")
    use reagentsPackage = new ExcelPackage (reagentsInfo)
    let myTools = reagentsPackage.Workbook.Worksheets.["tools"]

    //Documents
    let rqstform = "C:/Users/ipedroza/source/repos/FRM-M0206 Ghost Probe Synthesis Request.docx"
    let ligationForm = "C:/Users/ipedroza/source/repos/Formatted Forms/FRM-M0051-11_Ghost Probe Ligation using Excess Ghost Probe Oligo.docx"
    let gelForm = "C:/Users/ipedroza/source/repos/FRM-M0217-04_RUO Gel Electrophoresis QC for Ghost Probe Ligations Batch Record.docx"
    let gpZAG = "C:/Users/ipedroza/source/repos/FRM-10465-01_100-mer ZAG QC Batch Record.docx"
    let purificationForm = "C:/Users/ipedroza/source/repos/FRM-M0052-10 Purification of Ghost Probes with F-MODBs.docx"

    let user = Environment.UserName

    //Starts reading values of excel and stores it in "param"
   // try
    for param in inputParams do
        
        if processInput.Equals("ligate") then 
            Ligations.ligationStart inputParams rqstform ligationForm ghost myTools
               

        elif processInput.Equals("gelqc") then 
            GelQC.gelQcStart inputParams gelForm ghost myTools
         

        //elif processInput.Equals("zagqc") then

        //    Console.WriteLine "How many plates are you running?"
        //    let plateInput = Console.ReadLine() |> float

        //    let input2 = "zagqc"

        //    let gelVolume = (((plateInput - 1.0) * 5.0) + 20.0).ToString()
        //    let ureaVolume = (plateInput * 2.625).ToString()
        //    let diVolume = (plateInput * 315.0).ToString()
        //    let tenMerVolume = (plateInput * 60.0).ToString()

        //    for param in inputParams do
            
        //        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (csInfo param ghost)

        //        let gelLot = listFunction processInput myTools 2 
        //        let sybrLot = listFunction processInput myTools 3
        //        let ibLot = listFunction processInput myTools 4
        //        let ccLot = listFunction processInput myTools 5
        //        let ureaLot = listFunction processInput myTools 6
        //        let tenmerLot = listFunction processInput myTools 7
        //        let mpLot = listFunction processInput myTools 8 

        //        let docArray = File.ReadAllBytes(gpZAG)
        //        use docCopy = new MemoryStream(docArray)
        //        use zagDocument = WordprocessingDocument.Open (docCopy, true)
        //        let zagBody = zagDocument.MainDocumentPart.Document.Body

        //        let numberOfPlates = geneNumber / 96.0 
        //        let totalPlates = Math.Ceiling(numberOfPlates)
        //        let lotPlates = 
        //            if totalPlates = 1.0 then 
        //                (writeCsInfo zagBody 2 11).Text <- geneNumber.ToString()
        //                "1"

        //            else
        //                Console.WriteLine ("For " + param + " are you running the last plate?")
        //                let answer = Console.ReadLine()
        //                if answer = "yes" then
        //                    (writeCsInfo zagBody 2 11).Text <- geneNumber.ToString()
        //                    "1 - " + totalPlates.ToString()
        //                else       
        //                    (writeCsInfo zagBody 2 11).Text <- (geneNumber - (96.0 * (geneNumber / 96.0))).ToString()
        //                    "1 - " + (totalPlates - 1.0).ToString()
            

        //        //CS lot identifier info and number of plates being run
        //        (writeCsInfo zagBody 1 5).Text <- lot + " " + csName
        //        (writeCsInfo zagBody 1 13).Text <- lotPlates
        //        (writeCsInfo zagBody 1 21).Text <- totalPlates |> string
        //        (writeCsInfo zagBody 2 17).Text <- geneNumber |> string
        //        (writeCsInfo zagBody 2 21).Text <- scale |> string

        //        //finds cell of each reagent
        //        (fillCells zagBody 0 1 4 0 0).Text <- gelLot
        //        (fillCells zagBody 0 2 4 0 0).Text <- sybrLot
        //        (fillCells zagBody 0 3 4 0 0).Text <- ibLot
        //        (fillCells zagBody 0 4 4 0 0).Text <- ccLot
        //        (fillCells zagBody 0 5 4 0 0).Text <- ureaLot
        //        (fillCells zagBody 0 7 4 0 0).Text <- tenmerLot
        //        (fillCells zagBody 0 8 4 0 0).Text <- mpLot
        //        (fillCells zagBody 0 9 4 0 0).Text <- "N/A"

        //        //Calculations text and footnotes
        //        let gelCalculations = (fillCells zagBody 1 2 1 2 8)
        //        let gelFootNote = (footNoteSize zagBody 1 2 1 2 9)
        //        let sybrCalculations = (fillCells zagBody 1 2 1 4 6)
        //        let sybrFootNote = (footNoteSize zagBody 1 2 1 4 7)
        //        let ureaCalculations = (fillCells zagBody 1 7 1 2 2)
        //        let ureaFootNote = (footNoteSize zagBody 1 7 1 2 3)
        //        let diCalculations = (fillCells zagBody 1 7 1 4 5)
        //        let diFootNote = (footNoteSize zagBody 1 7 1 4 6)
        //        let tenMerCalculations = (fillCells zagBody 1 7 1 6 11)
        //        let tenMerFootNote = (footNoteSize zagBody 1 7 1 6 12)
            

        //        //Adds footnote to calculations section and comment section
        //        let firstLot = inputParams.[0]
        //        let restOfList = inputParams.[1..inputParams.Length-1]
        //        if inputParams.Length > 1 then       
        //            if param = firstLot then
        //                gelCalculations.Text <- gelVolume
        //                gelFootNote.Text <- "①"
        //                sybrCalculations.Text <- gelVolume
        //                sybrFootNote.Text <- "①"
        //                ureaCalculations.Text <- ureaVolume
        //                ureaFootNote.Text <- "①"
        //                diCalculations.Text <- diVolume
        //                diFootNote.Text <- "①"
        //                tenMerCalculations.Text <- tenMerVolume
        //                tenMerFootNote.Text <- "①"
        //            else
        //                gelFootNote.Text <- "①"
        //                sybrFootNote.Text <- "①"
        //                ureaFootNote.Text <- "①"
        //                diFootNote.Text <- "①"
        //                tenMerFootNote.Text <- "①"
        //            zagNote zagBody inputParams param
        //        else
        //            gelCalculations.Text <- gelVolume
        //            sybrCalculations.Text <- gelVolume
        //            ureaCalculations.Text <- ureaVolume
        //            diCalculations.Text <- diVolume
        //            tenMerCalculations.Text <- tenMerVolume
        //            gelFootNote.Text <- ""
        //            sybrFootNote.Text <- ""
        //            ureaFootNote.Text <- ""
        //            diFootNote.Text <- ""
        //            tenMerFootNote.Text <- ""

        //        let zagBatchRecordPath = "C:/Users/" + User + "/AppData/Local/Temp/ "+param + " Zag Batch Record" + ".docx"
        //        zagDocument.SaveAs(zagBatchRecordPath).Close()

        //elif processInput.Equals("purify") then
        //    //Console.WriteLine "Which reagents will you use?"
        //    let reagentsInput = "purifications" //Console.ReadLine ()

        //    Console.WriteLine "What is the regen of the beads being used?"
        //    let regenNumber = Console.ReadLine ()

        //    Console.WriteLine "At which bench are you working?"
        //    let benchId = Console.ReadLine ()

        //    for param in inputParams do
               
        //            //Takes value of each cell of the row in which the input lies and stores it for use in filling out Word Doc
        //            let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (csInfo param ghost)

        //            //Reading from reagents sheet and loking for reagent box input
        //            let sspeLot = listFunction  reagentsInput myTools  2
        //            let sspeExp = listFunction  reagentsInput myTools  3
        //            let sspeUbd = listFunction  reagentsInput myTools  4
        //            let tweenLot = listFunction  reagentsInput myTools  5
        //            let tweenExp = listFunction  reagentsInput myTools  6
        //            let tweenUbd = listFunction  reagentsInput myTools  7
        //            let depcLot = listFunction  reagentsInput myTools  8
        //            let depcExp = listFunction  reagentsInput myTools  9
        //            let depcUbd = listFunction  reagentsInput myTools  10
        //            let oneXLot = listFunction  reagentsInput myTools  11
        //            let oneXExp = listFunction  reagentsInput myTools  12
        //            let oneXUbd = listFunction  reagentsInput myTools  13
        //            let halfXLot = listFunction  reagentsInput myTools  14
        //            let halfXExp = listFunction  reagentsInput myTools  15
        //            let halfXUbd = listFunction  reagentsInput myTools  16
        //            let hybId = listFunction  reagentsInput myTools  17
        //            let HybCalibrationDate = listFunction  reagentsInput myTools  18
        //            let nanoDropId = listFunction  reagentsInput myTools  19
        //            let nanoDropCalibration = listFunction  reagentsInput myTools  20

        //            //Reading from reagents sheet and looking for bench input to find pipettes associated with that bench
        //            let pipetteCal = listFunction benchId myTools  2
        //            let p1000 = listFunction benchId myTools  3
        //            let p1000Id = listFunction benchId myTools  4
        //            let p200 = listFunction benchId myTools  5
        //            let p200Id = listFunction benchId myTools  6
        //            let p20 = listFunction benchId myTools  7
        //            let p20Id = listFunction benchId myTools  8
        //            let p2000 = listFunction benchId myTools  9
        //            let p2000Id = listFunction benchId myTools  10
        //            let p2 = listFunction benchId myTools  11
        //            let p2Id = listFunction benchId myTools  12

        //            //Reads in copy of Word Doc and starts processing
        //            let docArray = File.ReadAllBytes(purificationForm)
        //            use docCopy = new MemoryStream(docArray)
        //            use purificationDocument = WordprocessingDocument.Open (docCopy, true)
        //            let purificationBody = purificationDocument.MainDocumentPart.Document.Body

        //            //Fills out CS identifying info
        //            (writeCsInfo purificationBody 1 5).Text <- lot + " " + csName
        //            (writeCsInfo purificationBody 1 13).Text <- geneNumber |> string
        //            (writeCsInfo purificationBody 1 19).Text <- scale |> string

        //            //Fills reagent and equipment info 
        //            (purificationsFunction purificationBody 0 1 2 0 0 0).Text <- lot
        //            (purificationsFunction purificationBody 0 1 3 0 0 0).Text <- "N/A"
        //            (purificationsFunction purificationBody 0 2 2 0 0 0).Text <- sspeLot
        //            (purificationsFunction purificationBody 0 2 3 1 0 0).Text <- sspeExp
        //            (purificationsFunction purificationBody 0 2 3 2 3 0).Text <- sspeUbd
        //            (purificationsFunction purificationBody 0 3 2 0 0 0).Text <- tweenLot
        //            (purificationsFunction purificationBody 0 3 3 1 0 0).Text <- tweenExp
        //            (purificationsFunction purificationBody 0 3 3 2 3 0).Text <- tweenUbd
        //            (purificationsFunction purificationBody 0 4 2 0 0 0).Text <- depcLot
        //            (purificationsFunction purificationBody 0 4 3 1 0 0).Text <- depcExp
        //            (purificationsFunction purificationBody 0 4 3 2 3 0).Text <- depcUbd
        //            (purificationsFunction purificationBody 0 5 0 1 2 1).Text <- regenNumber
        //            (purificationsFunction purificationBody 0 5 2 0 0 0).Text <- regenNumber
        //            (purificationsFunction purificationBody 0 5 3 0 0 0).Text <- "N/A"
        //            (purificationsFunction purificationBody 0 6 2 0 0 0).Text <- oneXLot
        //            (purificationsFunction purificationBody 0 6 3 1 0 0).Text <- oneXExp
        //            (purificationsFunction purificationBody 0 6 3 2 3 0).Text <- oneXUbd
        //            (purificationsFunction purificationBody 0 7 2 0 0 0).Text <- halfXLot
        //            (purificationsFunction purificationBody 0 7 3 1 0 0).Text <- halfXExp
        //            (purificationsFunction purificationBody 0 7 3 2 3 0).Text <- halfXUbd
        //            (purificationsFunction purificationBody 0 9 1 0 0 0).Text <- hybId
        //            (purificationsFunction purificationBody 0 9 2 0 0 0).Text <- HybCalibrationDate
        //            (purificationsFunction purificationBody 0 10 1 0 0 0).Text <- nanoDropId
        //            (purificationsFunction purificationBody 0 10 2 0 0 0).Text <- nanoDropCalibration
        //            (purificationsFunction purificationBody 0 12 0 0 0 0).Text <- p1000
        //            (purificationsFunction purificationBody 0 12 1 0 1 0).Text <- p1000Id
        //            (purificationsFunction purificationBody 0 12 2 0 0 0).Text <- pipetteCal
        //            (purificationsFunction purificationBody 0 13 0 0 0 0).Text <- p200
        //            (purificationsFunction purificationBody 0 13 1 0 1 0).Text <- p200Id
        //            (purificationsFunction purificationBody 0 13 2 0 0 0).Text <- pipetteCal
        //            (purificationsFunction purificationBody 0 14 0 0 0 0).Text <- p20
        //            (purificationsFunction purificationBody 0 14 1 0 1 0).Text <- p20Id
        //            (purificationsFunction purificationBody 0 14 2 0 0 0).Text <- pipetteCal
        //            (purificationsFunction purificationBody 0 15 0 0 0 0).Text <- p2000
        //            (purificationsFunction purificationBody 0 15 1 0 1 0).Text <- p2000Id
        //            (purificationsFunction purificationBody 0 15 2 0 0 0).Text <- pipetteCal
        //            (purificationsFunction purificationBody 0 16 0 0 0 0).Text <- p2
        //            (purificationsFunction purificationBody 0 16 1 0 1 0).Text <- p2Id
        //            (purificationsFunction purificationBody 0 16 2 0 0 0).Text <- pipetteCal

               
        //            let volume = System.Math.Ceiling(theoreticalVolume scale geneNumber)
        //            let oneXVolume = System.Math.Round(((volume * 0.95) * 1.2), 1) 
        //            let sspeVolume = System.Math.Round((oneXVolume / 20.0), 1)
        //            let tweenVolume = System.Math.Round((oneXVolume / 100.0), 1)
        //            let water = System.Math.Round((oneXVolume - volume - sspeVolume - tweenVolume), 1)
        //            let calculatingBeadVolume = (oneXVolume * 2.0)


        //            let beads = 
        //                if(calculatingBeadVolume % 100.0) = 0.0 then
        //                    (oneXVolume * 2.0).ToString()
        //                else
        //                    ((100.0 - (calculatingBeadVolume % 100.0)) + calculatingBeadVolume).ToString()

        //            let buffer = 
        //                if geneNumber < 100.0 then 
        //                    (((((scale) - 1.0) * 1000.0) * 0.6) / (618.0 / (geneNumber)))
                       
        //                elif geneNumber < 400.0 then 
        //                    ((((((scale) - 1.0) * 1000.0) * 0.6)) / 5.0)
        //                else
        //                    ((((((scale) - 1.0) * 1000.0) * 0.6)) / 3.6)
        //            let elutionBuffer = (System.Math.Ceiling(buffer)).ToString()
               
        //            (purificationsFunction purificationBody 1 4 1 8 2 0).Text <- volume.ToString()
        //            (purificationsFunction purificationBody 1 5 1 0 1 0).Text <- oneXVolume.ToString()
        //            (purificationsFunction purificationBody 1 8 1 0 5 0).Text <- volume.ToString()
        //            (purificationsFunction purificationBody 1 8 1 1 6 0).Text <- sspeVolume.ToString()
        //            (purificationsFunction purificationBody 1 8 1 2 5 0).Text <- tweenVolume.ToString()
        //            (purificationsFunction purificationBody 1 8 1 3 2 0).Text <- water.ToString()
        //            (purificationsFunction purificationBody 1 8 1 5 3 0).Text <- oneXVolume.ToString()
        //            (purificationsFunction purificationBody 1 9 1 0 2 0).Text <- elutionBuffer
        //            (purificationsFunction purificationBody 1 10 1 1 2 0).Text <- beads
        //            (purificationsFunction purificationBody 1 13 1 0 2 0).Text <-beads
        //            (purificationsFunction purificationBody 1 19 1 2 2 0).Text <- beads
        //            (purificationsFunction purificationBody 1 22 1 2 2 0).Text <- beads
        //            (purificationsFunction purificationBody 1 25 1 2 2 0).Text <- elutionBuffer

        //            let purificationFormPath = "C:/Users/" + User + "/AppData/Local/Temp/ "+param + " Purification Form" + ".docx"
        //            purificationDocument.SaveAs(purificationFormPath).Close() |> ignore
    //with 
    //    | _ -> 
    //        SentryClientExtensions.CaptureMessage
    //        Exception() |> SentrySdk.CaptureException
    //        SentrySdk.AddBreadcrumb(inputParams.ToString())
    //        for param in inputParams do 
    //            let docs = pathsList user param

    //            for each in docs do 
    //                if File.Exists(each) then 
    //                    File.Delete(each)
    //                    Console.WriteLine "User Error..."
            
    //try
    //    for param in inputParams do 
    //        let docs = pathsList user param
    //        for each in docs do 
    //            if File.Exists(each) then 
    //                printDocuments each
    //with 
    //    | _ -> 
    //        Console.WriteLine "Unable to print documents"

    0 // return an integer exit code
    
