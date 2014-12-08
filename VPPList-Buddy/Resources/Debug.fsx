#I "../bin/Debug/"
#r "ExcelLibrary.dll"
#r "VPPList-Buddy.dll"

open System
open System.IO
open VPPListBuddy
open VPPListBuddy.VPP
open VPPListBuddy.IO
open VPPListBuddy.XLS
open VPPListBuddy.Workflow
open ExcelLibrary.SpreadSheet

let assetspath = __SOURCE_DIRECTORY__ 
let pathgoodvpp = Path.Combine(assetspath,"Sample.xls")
let pathbadtotals = Path.Combine(assetspath,"Incorrect Totals.xls")

let basexls = xlsopenfile pathgoodvpp

let randomvpp () = 
    let randomcodeentry = {Code=IO.Path.GetRandomFileName(); URL=IO.Path.GetRandomFileName()}
    blankspreadsheet
    |> fun vpp -> {vpp with VPPCodes = randomcodeentry :: vpp.VPPCodes}
    |> fun vpp -> {vpp with OrderID = "8675309"; Product="Best App v2.4";Purchaser="Fortune 500 Rep";CodesPurchased=1;CodesRedeemed=0;CodesRemaining=1;ProductType="Application";AdamID="Hmm?"}

let debugvppparser path = //Keep
    xlsopenfile path 
    |> Option.bind (xlsextractsheet' 0)
    |> Option.bind(xlsparsevpp) 

let openxls path =
    let openxls = new OpenXLSWorkFlow ()
    openxls.OnNonExistentFile <- new FileError(fun path -> printfn "File %s does not exist!" path)
    openxls.OnInvalidVPP <- new Alert(fun () -> printfn "VPP not formatted")

    match openxls.TryOpenVPP path with
    | true , partitioner -> printfn "good"
    | _ -> printfn "what happened?"

  
