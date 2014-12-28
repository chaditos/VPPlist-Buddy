#load "Debug NPOI.fsx"
#I "../bin/Release/"
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
let onoverallocation = new AllocationError(fun obj args -> printfn "Error: Over allocated by %d." args.Amount)
let onaddoverallocation = new EntryAllocationError(fun obj args -> printfn "Error: Over allocated by added %d." args.Amount)
let onduplicatename = new PartitionError(fun obj args -> printfn "Error: Name %s is repeated." args.PartitionEntry.Name )
let oninvalidfilename = new PartitionError(fun ob args -> printfn "Error: Name %s contains invalid characters." args.PartitionEntry.Name)
let onunderallocation =  new AllocationError(fun obj args -> printfn "Error: Under-allocated by %d" args.Amount)
let onnonexistentfile = new FileError(fun path -> printfn "Error: File at %s does not exist." path)
let onemptyworkbook = new FileError(fun path -> printfn "Error: XLS at %s seems to be empty." path)
let onparsefailure = new FileError(fun path -> printfn "Error: Application could not parse XLS at %s" path)
let oninvalidvpp = new Alert(fun () -> printfn "Error: XLS is not formatted as VPP Spreadsheet.")
let oninvalidpartitionsheet = new Alert(fun () -> printfn "Error: There is no partitionsheet")
let partionersetup = 
    new PartitionWorkflowSetup( fun pw ->
        pw.OnAddOverAllocation.AddHandler(onaddoverallocation)
        pw.OnAddDuplicateName.AddHandler(onduplicatename)
        pw.OnAddInvalidFileName.AddHandler(oninvalidfilename)
        pw.OnPartitionOverAllocation.AddHandler(onoverallocation)
        pw.OnPartitionUnderAllocation.AddHandler(onunderallocation)
    )
//let ondestinationerror = new FileError(fun path -> printfn "Error: Directory %s could not be written too." path)



let randomvpp () = 
    let randomcodeentry = {Code=IO.Path.GetRandomFileName(); URL=IO.Path.GetRandomFileName()}
    blankspreadsheet
    |> fun vpp -> {vpp with VPPCodes = randomcodeentry :: vpp.VPPCodes}
    |> fun vpp -> {vpp with OrderID = "8675309"; Product="Best App v2.4";Purchaser="Fortune 500 Rep";CodesPurchased=1;CodesRedeemed=0;CodesRemaining=1;ProductType="Application";AdamID="Hmm?"}

let debugvppparser path = //Keep
    xlsopenfile path 
    |> Option.bind (xlsextractsheet' 0)
    |> Option.bind(xlsparsevpp) 

let codes = [1..100] |> List.map(fun n -> {Code=n.ToString();URL=sprintf "apple.com/vpp/app/%d.html" n})
let testpartitioner() =
    let codeamount = List.length codes
    let vpp = {randomvpp() with VPPCodes = codes ; CodesRemaining = codeamount; CodesPurchased = codeamount; CodesRedeemed = 0}
    let pw = PartitionWorkflow(vpp)
    do 
        pw.OnAddOverAllocation.AddHandler(onaddoverallocation)
        pw.OnAddDuplicateName.AddHandler(onduplicatename)
        pw.OnAddInvalidFileName.AddHandler(oninvalidfilename)
        pw.OnPartitionOverAllocation.AddHandler(onoverallocation)
        pw.OnPartitionUnderAllocation.AddHandler(onunderallocation)
        pw.Add(new PartitionEntry("B128",101))
        pw.Add(new PartitionEntry("B129",1))
        pw.Add(new PartitionEntry("B12",30))
        pw.Add(new PartitionEntry("B12&*",30))
        pw.WriteToXLS(assetspath)
    
let openxls path =
    let openxls = new OpenXLSWorkFlow ()
    do
        openxls.OnEmptyWorkbook <- onemptyworkbook
        openxls.OnInvalidPartitionSheet <- oninvalidpartitionsheet
        openxls.OnInvalidVPP <- oninvalidvpp
        openxls.OnNonExistentFile <- onnonexistentfile
        openxls.OnParseFailure <- onparsefailure
    match openxls.TryOpenVPP(path,partionersetup) with
    | true , partitioner -> printfn "good"
    | _ -> () //errors should have been handled.

  
