namespace VPPListBuddy

open System
open System.IO
open VPPListBuddy.VPP
open VPPListBuddy.XLS

type public PartitionEntry(name:string,allocation:int) =
    let mutable name = name
    let mutable allocation = allocation

    let modified = new Event<PartitionEntry>()
    let deleted = new Event<PartitionEntry>()

    [<CLIEvent>]
    member this.Modified =  modified.Publish
    [<CLIEvent>]
    member this.Deleted =  deleted.Publish

    member this.Name
        with get() = name
        and set(value) = 
            name <- value
            modified.Trigger(this)
    member this.Allocation
        with get() = allocation
        and set(value) = 
            allocation <- value
            modified.Trigger(this)
    member this.Delete() =
        deleted.Trigger this


type AllocationError = delegate of int -> unit
type DirectoryChooser = delegate of unit -> string
type PartitionMaker = delegate of unit -> list<string * int>
type FileError = delegate of string -> unit
type PartitionError = delegate of PartitionEntry -> unit


type public PartitionWorkflow (VPP:VPP) = 

    let measure' (vpp:VPP) (partitions:seq<PartitionEntry>)  =
        let totalcodes = vpp.CodesRemaining
        let allocatedcodes = partitions |> Seq.sumBy(fun p -> p.Allocation)
        totalcodes - allocatedcodes

    let (|Over|Under|Exact|) allocation =
        match allocation with
        | _ when allocation < 0 -> Over
        | _ when allocation > 0 -> Under
        | _ -> Exact

    let partition (allocations:seq<PartitionEntry>) (vpp:VPP) =
        let rec allocatecodes acc sections codes  = //unchecked
            let partition entry =
                let rec helpertake acc n l = if n <= 0 then acc else helpertake (List.head l :: acc) (n-1) (List.tail l)
                fst entry , helpertake [] (snd entry) codes
            let remainingcodes c =
                let rec helperskip n l = if n <= 0 then l else helperskip (n-1) (List.tail l)
                helperskip (snd c) codes
            match sections with
            | [] -> acc
            | h :: remainingsections -> allocatecodes (partition h :: acc) remainingsections (remainingcodes h)
        let tovpps (vpp:VPP) pe =
            pe |> Seq.map( fun (name,codes) ->
                let numofcodes = Seq.length codes
                {vpp with 
                    FileName = name
                    VPPCodes = codes;
                    CodesRedeemed = 0;
                    CodesPurchased = numofcodes
                    CodesRemaining = numofcodes})
        let aslist (pe:seq<PartitionEntry>) = pe |> Seq.map(fun p -> p.Name , p.Allocation) |> Seq.toList
        let totalavailable = List.length vpp.AvailableVPPCodes
        let wantedamount = allocations |> Seq.sumBy(fun s -> s.Allocation)
        match totalavailable - wantedamount with
        | 0 -> Some(allocatecodes [] (aslist allocations) vpp.VPPCodes |> tovpps vpp)
        | _ -> None
    //change so you can modify extension
    let write directory vpps = vpps |> Seq.iter(fun vpp -> savevpptoxls vpp (Path.Combine(directory,vpp.FileName + ".xls")))
   
    let partitions = Collections.Generic.List<PartitionEntry>()

    //Events
    let partitionchanged = new Event<PartitionEntry>()//DelegateEvent<PartitionError>()
    let overallocation = new Event<int>()
    let onunderallocation = new Event<int>()
    let duplicatenames = new Event<PartitionEntry>()

    [<CLIEvent>] member this.OnDuplicateName = duplicatenames.Publish
    [<CLIEvent>] member this.OnOverAllocation = overallocation.Publish
    [<CLIEvent>] member this.OnUnderAllocation = onunderallocation.Publish

    member this.VPP = VPP

    member this.Add(entry:PartitionEntry) =
        let isduplicate = partitions |> Seq.exists(fun pe -> pe.Name = entry.Name) 
        let allocmeasure = measure' this.VPP partitions

        match isduplicate with
        | true -> duplicatenames.Trigger(entry)
        | false ->
            match allocmeasure with
            | Over as num -> overallocation.Trigger(num)
            | Under as num -> partitions.Add(entry)// onunderallocation.Trigger(num) do modified event
            | Exact -> partitions.Add(entry)

    member this.Clear() = 
        do
            partitions |> Seq.iter(fun p -> p.Delete())
            partitions.Clear()

    member this.EvenPartition(onzeropartitions,onleftover,oneven) =
        match Seq.length partitions with
        | 0 -> do onzeropartitions
        | _ -> 
            let leftover = this.VPP.CodesRemaining % Seq.length partitions
            let codeportion = this.VPP.CodesRemaining / Seq.length partitions
            do partitions |> Seq.iter(fun p -> p.Allocation <- codeportion)
            match leftover,codeportion with 
            | _,0 -> () //Can't partition with more entries than there is codes
            | 0,_ -> do partitions |> oneven
            | l,_ -> do partitions |> onleftover (PartitionEntry("",l))

    member this.Measure(onexact) =
        match measure' this.VPP partitions with
        | Over as num -> overallocation.Trigger(abs num)
        | Under as num -> onunderallocation.Trigger(abs num)
        | Exact -> do onexact
    
    member this.Partition(directorypath,onerror:FileError) =
        match measure' this.VPP partitions with
        | Over as num -> overallocation.Trigger(abs num)
        | Under as num -> onunderallocation.Trigger(abs num)
        | Exact -> 
            match partition partitions this.VPP with 
            | Some(vpps) -> write directorypath vpps 
            | None -> do onerror.Invoke(directorypath)
