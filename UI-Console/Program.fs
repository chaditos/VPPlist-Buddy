//For demonstrating the ability to create arbitrary UI's

open System
open VPPListBuddy.Workflow

[<EntryPoint>]
let main argv = 
    //Error message delegates and workflow setup
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
    let partitionersetup = 
        new PartitionWorkflowSetup( fun pw ->
            pw.OnAddOverAllocation.AddHandler(onaddoverallocation)
            pw.OnAddDuplicateName.AddHandler(onduplicatename)
            pw.OnAddInvalidFileName.AddHandler(oninvalidfilename)
            pw.OnPartitionOverAllocation.AddHandler(onoverallocation)
            pw.OnPartitionUnderAllocation.AddHandler(onunderallocation)
        )

    //Initialize Open XLS Workflow
    let openworkflow = new OpenXLSWorkFlow()
    do
        openworkflow.OnInvalidPartitionSheet <- oninvalidpartitionsheet
        openworkflow.OnNonExistentFile <- onnonexistentfile
        openworkflow.OnParseFailure <- onparsefailure
        openworkflow.OnEmptyWorkbook <- onemptyworkbook
        openworkflow.OnInvalidVPP <- oninvalidvpp

    //Try parsing xls files
    do 
        match Seq.length argv with
        | 2 ->
            let file,dest = Seq.head argv , Seq.last argv
            match openworkflow.TryOpenVPP(file,partitionersetup) with
            | true, partitioner -> 
                do 
                    partitioner.WriteToXLS(dest)
            | _ -> () 
        | _ -> printfn "Error: This program takes 2 arguments, a source file and destination directory."
    0