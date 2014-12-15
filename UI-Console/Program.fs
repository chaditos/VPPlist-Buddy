// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

open System
open VPPListBuddy.Workflow

[<EntryPoint>]

let main argv = 
    //Error message delegates
    let onnonexistentfile = new FileError(fun path -> printfn "Error: File at %s does not exist." path)
    let onxlsparsefailure = new FileError(fun path -> printfn "Error: Application could not parse XLS at %s" path)
    let onemptyworkbook = new FileError(fun path -> printfn "Error: XLS at %s seems to be empty." path)
    let oninvalidvpp = new Alert(fun () -> printfn "Error: XLS is not formatted as VPP Spreadsheet.")
    let onduplicatename = new PartitionError(fun (entry:PartitionEntry) -> printfn "Error: Name %s is repeated." entry.Name)
    let onunderallocation =  new AllocationError(fun num -> printfn "Error: Under-allocated by -%d" num)
    let onoverallocation = new AllocationError( fun num -> printfn "Error: Over-allocated by %d" num)
    let ondestinationerror = new FileError(fun path -> printfn "Error: Directory %s could not be written too." path)

    do 
        //Initialize Open XLS Workflow
        let openworkflow = new OpenXLSWorkFlow()
        openworkflow.OnNonExistentFile <- onnonexistentfile
        openworkflow.OnParseFailure <- onxlsparsefailure
        openworkflow.OnEmptyWorkbook <- onemptyworkbook
        openworkflow.OnInvalidVPP <- oninvalidvpp
   
        //Try parsing xls files
        match Seq.length argv with
        | 2 ->
            let file,dest = Seq.head argv , Seq.last argv
            match openworkflow.TryOpenVPP file with
            | true, partitioner -> 
                do 
                    partitioner.OnDuplicateName.Add(onduplicatename.Invoke)
                    partitioner.OnUnderAllocation.Add(onunderallocation.Invoke)
                    partitioner.OnOverAllocation.Add(onoverallocation.Invoke)
                    partitioner.Partition(dest,ondestinationerror)
                    printfn "Success: VPP file %s was partitioned to directory %s" file dest
            | _ -> () 
        | _ -> printfn "Error: This program takes 2 arguments, a source file and destination directory."
    0
     // return an integer exit code

