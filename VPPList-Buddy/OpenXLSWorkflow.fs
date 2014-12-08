namespace VPPListBuddy.Workflow

open System
open System.IO
open System.Runtime.InteropServices
open VPPListBuddy
open VPPListBuddy.VPP
open VPPListBuddy.IO
open VPPListBuddy.XLS

type Alert = delegate of unit -> unit
type VPPHandler = delegate of VPP -> unit


type OpenXLSWorkFlow() =
    //Default alert delegates that do nothing.
    let mutable onnonexistentfile = new FileError(fun path -> ())
    let mutable onparsefailure = new FileError(fun path -> ())
    let mutable onemptyworkbook = new FileError(fun path ->())
    let mutable oninvalidvpp = new Alert(fun path ->())
    let mutable oninvalidpartitionsheet = new Alert(fun () -> ())


    let checkinteropnull string' = match String.IsNullOrEmpty(string') with | false -> Some string' | true -> None //No log just fail.
    let openworkbook path =
        let filealert (alert:FileError) option = match option with | Some _ -> option | None -> (do alert.Invoke(path)); None
        checkinteropnull path
        |> Option.bind (fun path -> fileexists path |> filealert onnonexistentfile)
        |> Option.bind (fun path -> xlsopenfile path |> filealert onparsefailure)
        |> Option.bind (fun workbook -> xlstestempty workbook |> filealert onemptyworkbook)

    let openworksheet = xlsextractsheet'

    let openworksheet' = xlsextractsheet

    let parsevppsheet = xlsparsevpp >> (fun vpp -> match vpp with | Some _ -> vpp | None -> (do oninvalidvpp.Invoke()); None)

    let parsepartitionsheet vppspreadsheet =
        let read, readint = vppspreadsheet |> xlstrystring, vppspreadsheet |> xlstryint
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

    //Setable Error Warning Delegates
    member this.OnNonExistentFile 
        with get() = onnonexistentfile
        and set(value) = onnonexistentfile <- value
    member this.OnParseFailure
        with get() = onparsefailure
        and set(value) = onparsefailure <- value
    member this.OnEmptyWorkbook
        with get() = onemptyworkbook
        and set(value) = onemptyworkbook <- value
    member this.OnInvalidVPP
        with get() = oninvalidvpp
        and set(value) = oninvalidvpp <- value
    member this.OnInvalidPartitionSheet 
        with get() = oninvalidpartitionsheet
        and set(value) = oninvalidpartitionsheet <- value

    member this.TryOpenVPP(path,[<Out>] partitionworkflow:byref<PartitionWorkflow>) =
        let workbook = openworkbook path
        let vppsheet = workbook |> Option.bind(openworksheet 0) |> Option.bind(parsevppsheet)
        let partitionsheet = workbook |> Option.bind(openworksheet 1) |> Option.bind(parsepartitionsheet)
        match vppsheet , partitionsheet with
        | Some(vpp) , None -> 
            do oninvalidpartitionsheet.Invoke ()
            false
        | Some(vpp) , Some(partitions) -> 
            let newpw = new PartitionWorkflow(vpp)
            do
                partitions
                |> Seq.iter(fun entry -> (newpw.Add( new PartitionEntry(fst entry, snd entry)) ))
                partitionworkflow <- newpw
            true
        | _ , _ -> false //Alerts would have been handled already i invalid

(*
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
*)
