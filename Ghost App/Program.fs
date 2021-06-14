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


let pathsList (user : string) (param : string) = 
    [
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Request Form" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Ligation Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Gel Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Zag Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " reQC Batch Record" + ".docx"
    "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Purification Form" + ".docx"
    ]


let printDocuments (path : string) (user : string) =
    let printing = new Process()
    printing.StartInfo.FileName <- path
    printing.StartInfo.WorkingDirectory <- "C:/Users/" + user + "/AppData/Local/Temp"
    printing.StartInfo.CreateNoWindow <- true
    printing.StartInfo.WindowStyle <- ProcessWindowStyle.Hidden
    printing.StartInfo.UseShellExecute <- true
    printing.EnableRaisingEvents <- false
    printing.StartInfo.Verb <- "Print"
    printing.Start() |> ignore

    while (printing.HasExited = false) do
        printing.WaitForExit(10000)
    File.Delete(path)
    
    

[<EntryPoint>]
let main argv =

    //Reading in exel doc
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial  
    
    //User interface - input will take string as input
    Console.WriteLine "What GP will you work on?"
    let input = Console.ReadLine ()
    let inputSplit = input.Split(' ')
    let inputParams = [for i in inputSplit do i.ToUpper()]

    Console.WriteLine "Which process are you conducting?"
    let processInput = Console.ReadLine ()

    use __ = SentrySdk.Init ( fun o ->
           o.Dsn <-  "https://e3c4cd9eb460410e89402ea524bb9922@o811036.ingest.sentry.io/5805218"
           o.SendDefaultPii <- true
           o.AttachStacktrace <- true
           o.ShutdownTimeout <- TimeSpan.FromSeconds 10.0 
           o.MaxBreadcrumbs <- 50 
           )

    SentrySdk.ConfigureScope(fun scope -> scope.SetTag("User Input", input) )
    SentrySdk.AddBreadcrumb(input)
    SentrySdk.ConfigureScope(fun newTag -> newTag.SetTag("Manufacturing_Process", processInput))

    
    
    //Reading in excel file from path on Reporter Probe sheet
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let fileInfo = new FileInfo("W:/Production/Reporter requests/CODESET TRACKERS/CodeSet Status.xlsx")
    use package = new ExcelPackage (fileInfo)
    let ghost = package.Workbook.Worksheets.["Upstream - GP"]

    //Reading in excel file from path on oligo sheet
    ExcelPackage.LicenseContext <- Nullable LicenseContext.NonCommercial
    let oligoInfo = new FileInfo("W:/Production/Probe Oligos/REMP Files/_Re-Rack Files/Rerack Status.xlsx")
    use oligoPackage = new ExcelPackage (oligoInfo)
    let oligoStamps = oligoPackage.Workbook.Worksheets.["CodeSet Archive"]

    //Reading in reagents excel book
    let reagentsInfo = new FileInfo("S:/ip/reagentsandtools.xlsx")
    use reagentsPackage = new ExcelPackage (reagentsInfo)
    let myTools = reagentsPackage.Workbook.Worksheets.["tools"]

    //Documents
    let rqstform = "W:/program_files/FRM-M0206 Ghost Probe Synthesis Request.docx"
    let ligationForm = "W:/program_files/FRM-M0051-11_Ghost Probe Ligation using Excess Ghost Probe Oligo.docx"
    let gelForm = "W:/program_files/FRM-M0217-04_RUO Gel Electrophoresis QC for Ghost Probe Ligations Batch Record.docx"
    let zagForm = "W:/program_files/FRM-10465-01_100-mer ZAG QC Batch Record.docx"
    let reQcForm = "W:/program_files/FRM-M0183-03_Ghost Probe Re-QC.docx"
    let purificationForm = "W:/program_files/FRM-M0052-10 Purification of Ghost Probes with F-MODBs.docx"

    let user = Environment.UserName


    //Starts reading values of excel and stores it in "param"
    try
        
        if processInput.Equals("ligate", StringComparison.InvariantCultureIgnoreCase) then 
            Ligations.ligationStart inputParams rqstform ligationForm ghost oligoStamps myTools
               

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
            
    try
        for param in inputParams do 
            let docs = pathsList user param
            for each in docs do 
                if File.Exists(each) then 
                    printDocuments each user
                else 
                    ignore()
        
                    
    with 
        | _ -> 
            Console.WriteLine "Unable to print documents"
            for param in inputParams do 
                let docs = pathsList user param
                for each in docs do 
                    if File.Exists(each) then
                        File.Delete(each)


    0 // return an integer exit code
    
