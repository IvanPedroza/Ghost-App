module Ligations


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


let ligationStart (inputParams : string list) (rqstForm : string) (ligationForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    let user =  Environment.UserName 
    let mutable myList = []   
    
    for param in inputParams do
        let reagentsInput = "gpligations"

        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = codesetIdentifiers param ghost

        //Reads in Word Doc
        let docArray = File.ReadAllBytes(rqstForm)
        let lengthDox = docArray.Length
        use _copyDoc = new MemoryStream(docArray)
        use rqstDocument = WordprocessingDocument.Open(_copyDoc, true)
      
        let rqstbody = rqstDocument.MainDocumentPart.Document.Body

        let codesetType = HelperFunctions.determineFormulation csName formulation
       
        //Formats the shipping date 
        let formattedShipDate =
            match shipDate with 
                | "TBD"  -> "TBD"
                | _ -> 
                    let firststring = shipDate.Substring(0,3)
                    let secondstring = shipDate.Substring(3,2).ToLower()
                    firststring + secondstring + (DateTime.Now.Year.ToString())

        //Finds text for GP lot in word doc
        (requestForms rqstbody 0 1 0 0).Text <- lot
        (requestForms rqstbody 0 1 1 0).Text <- csName
        (requestForms rqstbody 0 1 2 0).Text <- species
        (requestForms rqstbody 0 1 3 0).Text <- customer
        (requestForms rqstbody 0 1 4 0).Text <- geneNumber
        (requestForms rqstbody 0 1 5 0).Text <- scale
        (requestForms rqstbody 0 1 6 0).Text <- formattedShipDate
        (rqstFormDropdowns rqstbody 1 0 0 0 0).Text <- codesetType
        (rqstFormDropdowns rqstbody  2 0 0 0 0).Text <- formulation
   

        //used to replace checked box text
        let concentrationCheck = "☒" 
        (rqstFormDropdowns rqstbody 12 0 0 0 0).Text <- concentrationCheck

        let rqstFormPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Request Form" + ".docx"
        rqstDocument.SaveAs(rqstFormPath).Close() |> ignore
           
        //Same as above but for reagents excel doc
        let GlLot = ligationsListFunction reagentsInput myTools 2 
        let GlExp = ligationsListFunction reagentsInput myTools 3
        let T4BufferLot = ligationsListFunction reagentsInput myTools 4
        let T4BufferExpiration = ligationsListFunction reagentsInput myTools 5
        let BF2Lot = ligationsListFunction reagentsInput myTools 6
        let BF2Expiration = ligationsListFunction reagentsInput myTools 7
        let ATPLot = ligationsListFunction reagentsInput myTools 8 
        let ATPExpiration = ligationsListFunction reagentsInput myTools 9
        let H2OLot = ligationsListFunction reagentsInput myTools 10
        let H2OExp = ligationsListFunction reagentsInput myTools 11
        let H2OUbd = ligationsListFunction reagentsInput myTools 12
        let t4EnzymeLot = ligationsListFunction reagentsInput myTools 13
        let t4EnzymeExp = ligationsListFunction reagentsInput myTools 14

        //Reads in Word Doc and starts processing a copy of it
        let docArray = File.ReadAllBytes(ligationForm)
        use _copyDoc = new MemoryStream(docArray)
        use ligationDocument = WordprocessingDocument.Open(_copyDoc, true)
        let ligationsBody = ligationDocument.MainDocumentPart.Document.Body

        //Find text in the docx table and assigns it string values
        (ligationsCsInfoHeader ligationsBody 2 2).Text <- lot + " " + csName
        (ligationsCsInfoHeader ligationsBody 2 9).Text <- geneNumber
        (ligationsCsInfoHeader ligationsBody 2 16).Text <- scale
        (ligationsTableFiller ligationsBody 0 2 2 0 0).Text <- ""
        (ligationsTableFiller ligationsBody 0 2 3 0 0).Text <- "N/A"
        (ligationsTableFiller ligationsBody 0 3 2 0 0).Text <- GlLot
        (ligationsTableFiller ligationsBody 0 3 3 0 0).Text <- GlExp
        (ligationsTableFiller ligationsBody 0 4 2 0 0).Text <- T4BufferLot
        (ligationsTableFiller ligationsBody 0 4 3 0 0).Text <- T4BufferExpiration
        (ligationsTableFiller ligationsBody 0 5 2 0 0).Text <- BF2Lot
        (ligationsTableFiller ligationsBody 0 5 3 0 0).Text <- BF2Expiration
        (ligationsTableFiller ligationsBody 0 6 2 0 0).Text <- ATPLot
        (ligationsTableFiller ligationsBody 0 6 3 0 0).Text <- ATPExpiration
        (ligationsTableFiller ligationsBody 0 7 3 2 2).Text <- H2OUbd
        (ligationsTableFiller ligationsBody 0 7 2 0 0).Text <- H2OLot
        (ligationsTableFiller ligationsBody 0 7 3 1 1).Text <- H2OExp
        (ligationsTableFiller ligationsBody 0 8 2 0 0).Text <- t4EnzymeLot
        (ligationsTableFiller ligationsBody 0 8 3 0 0).Text <- t4EnzymeExp
              
        //calculates reagent amounts and assigns calculations to the last parameter for reference
        let lastLot = inputParams.Last()
        let lastLotScale = ligationsListFunction lastLot ghost 7
              
        if not (scale = lastLotScale) then 
            let scaling = ((scale |> float) / (lastLotScale |> float))
            let scaledGenes = (geneNumber |> float) * scaling
            myList <- scaledGenes :: myList
        else 
            myList <- (geneNumber |> float) :: myList

        let plateTotal = System.Math.Ceiling((geneNumber |> float) / 96.0) |> int
        let iterator = [1..9]
        let mutable naList = []
        for i in iterator do 
            if i <= plateTotal then 
                let x = ""
                naList <- x :: naList
            else 
                let x = "N/A"
                naList <- x :: naList

        (ligationsTableFiller ligationsBody 0 11 1 0 1).Text <- naList.[7]
        (ligationsTableFiller ligationsBody 0 11 2 0 0).Text <- naList.[7]
        (ligationsTableFiller ligationsBody 0 12 1 0 1).Text <- naList.[6]
        (ligationsTableFiller ligationsBody 0 12 2 0 0).Text <- naList.[6]
        (ligationsTableFiller ligationsBody 0 13 1 0 1).Text <- naList.[5]
        (ligationsTableFiller ligationsBody 0 13 2 0 0).Text <- naList.[5]
        (ligationsTableFiller ligationsBody 0 14 1 0 1).Text <- naList.[4]
        (ligationsTableFiller ligationsBody 0 14 2 0 0).Text <- naList.[4]
        (ligationsTableFiller ligationsBody 0 15 1 0 1).Text <- naList.[3]
        (ligationsTableFiller ligationsBody 0 15 2 0 0).Text <- naList.[3]
        (ligationsTableFiller ligationsBody 0 16 1 0 1).Text <- naList.[2]
        (ligationsTableFiller ligationsBody 0 16 2 0 0).Text <- naList.[2]
        (ligationsTableFiller ligationsBody 0 17 1 0 1).Text <- naList.[1]
        (ligationsTableFiller ligationsBody 0 17 2 0 0).Text <- naList.[1]
        (ligationsTableFiller ligationsBody 0 18 1 0 1).Text <- naList.[0]
        (ligationsTableFiller ligationsBody 0 18 2 0 0).Text <- naList.[0]

        let oligo, ligator, buffer, bf2, atp, water, ligase, masterMix =   oligoStamp (scale |> float )

        if param = lastLot then 
            let reactions = roundupbyfive (myList.Sum() * 1.1)
            let ligatorAdded = reactions * ligator
            let bufferAdded = reactions * buffer 
            let bf2Added = reactions * bf2
            let atpAdded = System.Math.Round((reactions * atp), 2)
            let waterAdded = reactions * water
            let ligaseAdded = reactions * ligase
            let atpUndiluted = roundupbyfive(((atpAdded * 15.0) / 100.0))
            let atpTotal = System.Math.Round(((atpUndiluted * 100.0) / 15.0), 1)
            let atpDilutant = System.Math.Round((atpTotal - atpUndiluted), 1)
            let aliquots = (ligatorAdded + bufferAdded + bf2Added + atpAdded + waterAdded + ligaseAdded) / 8.0
            (ligationsTableFiller ligationsBody 0 30 1 7 1).Text <- atpUndiluted.ToString()
            (ligationsTableFiller ligationsBody 0 30 1 7 9).Text <- atpTotal.ToString()
            (ligationsTableFiller ligationsBody 0 30 1 10 1).Text <- atpTotal.ToString()
            (ligationsTableFiller ligationsBody 0 30 1 10 5).Text <- atpUndiluted.ToString()
            (ligationsTableFiller ligationsBody 0 30 1 10 9).Text <- atpDilutant.ToString()

            (ligationsTableFiller ligationsBody 0 31 1 0 3).Text <- reactions.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 1 3).Text <- ligatorAdded.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 2 3).Text <- bufferAdded.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 3 4).Text <- bf2Added.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 4 2).Text <- atpAdded.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 5 2).Text <- waterAdded.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 6 2).Text <- ligaseAdded.ToString()
            (ligationsTableFiller ligationsBody 0 31 1 8 4).Text <- aliquots.ToString()
                
        (ligationsTableFiller ligationsBody 0 26 1 2 2).Text <- oligo
        (ligationsTableFiller ligationsBody 0 33 1 3 2).Text <- masterMix
        footnotes ligationsBody inputParams param

       
        let ligationBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Ligation Batch Record" + ".docx"
        ligationDocument.SaveAs(ligationBatchRecordPath).Close() |> ignore
