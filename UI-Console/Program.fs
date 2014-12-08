// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

open VPPListBuddy.Workflow

[<EntryPoint>]

let main argv = 
    //Alert delegates
    let onnonexistentfile = new FileError(fun path -> printfn "Error: File at %s does not exist." path)
    let onemptyworkbook = new FileError(fun path -> printfn "Error: XLS at %s seems to be empty." path)
    let onxlsparsefailure = new FileError(fun path -> printfn "Error: Application could not parse XLS at %s" path)
    let oninvalidvpp = new Alert(fun () -> printfn "Error: XLS is not formatted as VPP Spreadsheet.")
    let onduplicatename = new PartitionError(fun (entry:PartitionEntry) -> printfn "Error: Name %s is repeated." entry.Name)
    let onunderallocation =  new AllocationError(fun num -> printfn "Error: Under-allocated by -%d" num)
    let onoverallocation = new AllocationError( fun num -> printfn "Error: Over-allocated y %d" num)
    let ondestinationerror = new FileError(fun path -> printfn "Error: Directory %s could not be written too." path)

    //Initialize Open XLS Workflow
    let openworkflow = new OpenXLSWorkFlow()
    do 
        openworkflow.OnNonExistentFile <- onnonexistentfile
        openworkflow.OnParseFailure <- onxlsparsefailure
        openworkflow.OnEmptyWorkbook <- onemptyworkbook
        openworkflow.OnInvalidVPP <- oninvalidvpp

    //Try parsing xls files
    do
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
            | _ -> ()
        | _ -> ()
    0
     // return an integer exit code

