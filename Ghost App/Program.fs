﻿// Learn more about F# at http://fsharp.org

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open System.Diagnostics
open Sentry
open HelperFunctions

// Function used to save the filled out forms to temp folder for printing before deleting
let pathsList (user : string) (param : string) = 
    [
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Request Form" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Ligation Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Gel Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Zag Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Purification Form" + ".docx"
    ]

// Function for printing Batch Record to default printer
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

// Program entry point
[<EntryPoint>]
let main argv =

    // Function reading in Excel document form LIMS
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial  
    
    // User interface - input will take string as input
    Console.WriteLine "What GP will you work on?"
    let input = Console.ReadLine ()
    let inputSplit = input.Split(' ')
    let inputParams = [for i in inputSplit do i.ToUpper()]

    Console.WriteLine "Which process are you conducting?"
    let processInput = Console.ReadLine ()

    // Error logging using SentryIO
    use __ = SentrySdk.Init ( fun o ->
           o.Dsn <-  "https://d1553cf78c164e5d9813ca11cc417d80@o561151.ingest.sentry.io/5697684"
           o.SendDefaultPii <- true
           o.AttachStacktrace <- true
           o.ShutdownTimeout <- TimeSpan.FromSeconds 10.0 
           o.MaxBreadcrumbs <- 50 
           )

    SentrySdk.ConfigureScope(fun scope -> scope.SetTag("User Input", input) )
    SentrySdk.AddBreadcrumb(input)
    SentrySdk.ConfigureScope(fun newTag -> newTag.SetTag("Manufacturing_Process", processInput))

    
    
    // Funtion reading in Excel file from path on Reporter Probe sheet of LIMS
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let fileInfo = new FileInfo("W:/Production/Reporter requests/CODESET TRACKERS/CodeSet Status.xlsx")
    use package = new ExcelPackage (fileInfo)
    let ghost = package.Workbook.Worksheets.["Upstream - GP"]

    // Funtion reading in reagents Excel sheet of LIMS
    let reagentsInfo = new FileInfo("S:/ip/reagentsandtools.xlsx")
    use reagentsPackage = new ExcelPackage (reagentsInfo)
    let myTools = reagentsPackage.Workbook.Worksheets.["tools"]

    // Blank Batch Record template forms for filling
    let rqstform = "C:/Users/ipedroza/source/repos/FRM-M0206 Ghost Probe Synthesis Request.docx"
    let ligationForm = "C:/Users/ipedroza/source/repos/Formatted Forms/FRM-M0051-11_Ghost Probe Ligation using Excess Ghost Probe Oligo.docx"
    let gelForm = "C:/Users/ipedroza/source/repos/FRM-M0217-04_RUO Gel Electrophoresis QC for Ghost Probe Ligations Batch Record.docx"
    let zagForm = "C:/Users/ipedroza/source/repos/FRM-10465-01_100-mer ZAG QC Batch Record.docx"
    let reQcForm = "C:/Users/ipedroza/source/repos/FRM-M0183-03_Ghost Probe Re-QC.docx"
    let purificationForm = "C:/Users/ipedroza/source/repos/FRM-M0052-10 Purification of Ghost Probes with F-MODBs.docx"

    // Logged in user -- used for saving to temp folder and error logging
    let user = Environment.UserName


    // Starts reading values of Excel and stores it in "param"
    // If error occurs, it is logged on Sentry IO
    try
        
        if processInput.Equals("ligate", StringComparison.InvariantCultureIgnoreCase) then 
            Ligations.ligationStart inputParams rqstform ligationForm ghost myTools
               

        elif processInput.Equals("gelqc", StringComparison.InvariantCultureIgnoreCase) then 
            GelQC.gelQcStart inputParams gelForm ghost myTools
         

        elif processInput.Equals("zagqc", StringComparison.InvariantCultureIgnoreCase) then
            ZagQC.zagStart inputParams zagForm ghost myTools

        elif processInput.Equals("reqc", StringComparison.InvariantCultureIgnoreCase) then 
            ReQC.reQcStart inputParams reQcForm ghost myTools


        elif processInput.Equals("purify", StringComparison.InvariantCultureIgnoreCase) then
            Purifications.purificationStart inputParams purificationForm ghost myTools

        else 
            Console.WriteLine "Invalid Process Entry..."


    
    with 
        | ex ->
            ex |> SentrySdk.CaptureException |> ignore


           
            for param in inputParams do 
                let docs = pathsList user param

                for each in docs do 
                    if File.Exists(each) then 
                        File.Delete(each)
                        Console.WriteLine "User Error..."

    // Tries to print document to default printer -- notifies user when error occurs
    try
        for param in inputParams do 
            let docs = pathsList user param
            for each in docs do 
                if File.Exists(each) then 
                    printDocuments each
    with 
        | _ -> 
            Console.WriteLine "Unable to print documents"

    0 // return an integer exit code
    
