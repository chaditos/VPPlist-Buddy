#load "Debug NPOI.fsx"
#I "../bin/Debug/"
#r "VPPList-Buddy.dll"

open System
open System.IO
open VPPListBuddy
open VPPListBuddy.VPP
open VPPListBuddy.IO
open VPPListBuddy.XLS
open VPPListBuddy.Workflow


let assetspath = __SOURCE_DIRECTORY__ 
let pathgoodvpp = Path.Combine(assetspath,"Sample.xls")
let pathbadtotals = Path.Combine(assetspath,"Incorrect Totals.xls")
let pathxlstemplate = Path.Combine(assetspath,"Template.xls")
let basexls = xlsopenfile pathgoodvpp

//Alert Delegates
let onnonexistentfile = new FileError(fun path -> printfn "Error: File at %s does not exist." path)
let onemptyworkbook = new FileError(fun path -> printfn "Error: XLS at %s seems to be empty." path)
let onparsefailure = new FileError(fun path -> printfn "Error: Application could not parse XLS at %s" path)
let oninvalidvpp = new Alert(fun () -> printfn "Error: XLS is not formatted as VPP Spreadsheet.")
let onduplicatename = new PartitionError(fun (entry:PartitionEntry) -> printfn "Error: Name %s is repeated." entry.Name)
let onunderallocation =  new AllocationError(fun num -> printfn "Error: Under-allocated by -%d" num)
let onoverallocation = new AllocationError( fun num -> printfn "Error: Over-allocated y %d" num)
let ondestinationerror = new FileError(fun path -> printfn "Error: Directory %s could not be written too." path)



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
    openxls.OnNonExistentFile <- onnonexistentfile
    openxls.OnInvalidVPP <- oninvalidvpp
    openxls.OnParseFailure <- onparsefailure
    openxls.OnEmptyWorkbook <- onemptyworkbook
    //openxls.OnInvalidPartitionSheet <- on

    match openxls.TryOpenVPP path with
    | true , partitioner -> printfn "good"
    | _ -> printfn "what happened?"


let testpartitioner() =
    let codes = [1..100] |> List.map(fun n -> {Code=n.ToString();URL=sprintf "apple.com/vpp/app/%d.html" n})
    let vpp = {blankspreadsheet with VPPCodes = codes ; CodesRemaining = 100; CodesPurchased = 100; CodesRedeemed = 0}
    let pw = PartitionWorkflow(vpp)
    do 
        pw.OnDuplicateName.Add(onduplicatename.Invoke)
        pw.OnOverAllocation.Add(onoverallocation.Invoke)
        pw.OnUnderAllocation.Add(onunderallocation.Invoke)
        pw.Add(new PartitionEntry("B128",90))
        pw.Add(new PartitionEntry("B129",10))
        pw.Partition(assetspath,fun path -> printfn "Could not write to %s" path)

  
