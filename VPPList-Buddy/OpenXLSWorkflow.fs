namespace VPPListBuddy

open System
open System.IO
open VPPListBuddy
open VPPListBuddy.VPP
open VPPListBuddy.IO
open VPPListBuddy.XLS

type FileError = delegate of string -> unit
type Alert = delegate of unit -> unit
type VPPHandler = delegate of VPP -> unit
type PartitionError = delegate of PartitionEntry -> unit


type OpenXLSWorkFlow(OnNonExistentFile:FileError,OnParseFailure:FileError,OnEmptyWorkbook:FileError,OnInvalidVPP:Alert) =
    let checkinteropnull string' = match String.IsNullOrEmpty(string') with | false -> Some string' | true -> None //No log just fail.
    let openworkbook path =
        let filealert (alert:FileError) option = match option with | Some _ -> option | None -> (do alert.Invoke(path)); None
        checkinteropnull path
        |> Option.bind (fun path -> fileexists path |> filealert OnNonExistentFile)
        |> Option.bind (fun path -> xlsopenfile path |> filealert OnParseFailure)
        |> Option.bind (fun workbook -> xlstestempty workbook |> filealert OnEmptyWorkbook)

    let openworksheet = xlsextractsheet'

    let openworksheet' = xlsextractsheet

    let parsevppsheet = xlsparsevpp >> (fun vpp -> match vpp with | Some _ -> vpp | None -> (do OnInvalidVPP.Invoke()); None)

    let parsepartitionsheet vppspreadsheet =
        let read, readint = vppspreadsheet |> xlscellreader, vppspreadsheet |> xlscellintreader
        let rec readlines start acc =
            (start, 0, start, 1)
            |> fun (v1, h1, v2, h2) ->
                match read (v1,h1) , readint (v2,h2) with
                | None, Some(allocation:int) -> None
                | Some(partition), None -> None
                | Some(partition:string),Some(allocation:int) -> readlines (start + 1) ((partition,allocation) :: acc)
                | None, None when start = 0 -> None
                | _ -> Some(acc |> List.rev)
        readlines 0 []

    member this.OpenVPP(path,vpphandler:VPPHandler) =
        let vpp =
            openworkbook path
            |> Option.bind(openworksheet 0) 
            |> Option.bind(parsevppsheet)
        match vpp with
        | Some vpp -> vpphandler.Invoke(vpp)
        | None -> () //Alerts would have ben handled already if invalid.

    member this.OpenVPP(path,vpphandler:VPPHandler,partitioner) =
        let workbook = openworkbook path
        let vppsheet = workbook |> Option.bind(openworksheet 0) |> Option.bind(parsevppsheet)
        let partitionsheet = workbook |> Option.bind(openworksheet 1) |> Option.bind(parsepartitionsheet)
        match vppsheet , partitionsheet with
        | Some(vpp) , None -> vpphandler.Invoke(vpp)
        | Some(vpp) , Some(partition) -> (vpp,partition) ||> partitioner
        | _ , _ -> () //Alerts would have been handled already i invalid

    member this.WriteVPP(path,(onoverallocation:AllocationError),(onunderallocation:AllocationError),(onduplicatename:PartitionError),oninvalidpartitionsheet:Alert,directorychooser) =
        let workbook = openworkbook path
        let vppsheet = workbook |> Option.bind(openworksheet 0) |> Option.bind(parsevppsheet)
        let partitionsheet = workbook |> Option.bind(openworksheet 1) |> Option.bind(parsepartitionsheet)
        match vppsheet , partitionsheet with
        | Some(vpp) , None -> do oninvalidpartitionsheet.Invoke ()
        | Some(vpp) , Some(partition) -> 
            let partitioner = new PartitionWorkflow(vpp)
            do
                partitioner.OnOverAllocation.Add(onoverallocation.Invoke)
                partitioner.OnDuplicateName.Add(onduplicatename.Invoke)
                partitioner.OnUnderAllocation.Add(onunderallocation.Invoke)

                partition
                |> Seq.iter(fun entry ->  partitioner.Add( new PartitionEntry(fst entry, snd entry) ))

                partitioner.Partition(directorychooser , fun () -> printfn "Didn't work out")
        | _ , _ -> () //Alerts would have been handled already i invalid