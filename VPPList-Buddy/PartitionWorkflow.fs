namespace VPPListBuddy.Workflow

open System
open System.IO
open VPPListBuddy.VPP
open VPPListBuddy.XLS

type private AllocationMeasurement = //Can't put inside PartitionWorkflow class?
    | Over of Amount:int
    | Under of Amount:int
    | Exact

type public PartitionEntry(name:string,allocation:int) =
    let mutable name = name
    let mutable allocation = allocation
    member this.Name
        with get() = name
        and set(value) = name <- value
    member this.Allocation
        with get() = allocation
        and set(value) = allocation <- value

type DirectoryChooser = delegate of unit -> string
//.Net Friendly Events
type AllocationErrorEventArgs(Amount:int) =
    inherit EventArgs()
    member this.Amount = Amount
type AllocationError = delegate of obj * AllocationErrorEventArgs -> unit

type EntryAllocationErrorEventArgs(Amount:int,Entry:PartitionEntry) =
    inherit EventArgs()
    member this.Amount = Amount
    member this.Entry = Entry
type EntryAllocationError = delegate of obj * EntryAllocationErrorEventArgs -> unit

type PartitionErrorEventArgs(Entry:PartitionEntry) =
    inherit EventArgs()
    member this.PartitionEntry = Entry
type PartitionError = delegate of obj * PartitionErrorEventArgs -> unit

//Workflow 
type public PartitionWorkflow (VPP:VPP) = 

    let onaddduplicatename = new Event<PartitionError,PartitionErrorEventArgs>()
    let onaddinvalidfilename = new Event<PartitionError,PartitionErrorEventArgs>()
    let onaddoverallocation = new Event<EntryAllocationError,EntryAllocationErrorEventArgs>()
    let onpartitionunderallocation = new Event<AllocationError,AllocationErrorEventArgs>()
    let onpartitionoverallocation = new Event<AllocationError,AllocationErrorEventArgs>()

    let partitions = Collections.Generic.List<PartitionEntry>()
    let invaliddirchars = Set.ofArray (Path.GetInvalidPathChars())
    let invalidfilechars = Set.ofArray (Path.GetInvalidFileNameChars())

    let measure' (vpp:VPP) (partitions:seq<PartitionEntry>)  =
        let totalcodes = vpp.CodesRemaining
        let allocatedcodes = partitions |> Seq.sumBy(fun p -> p.Allocation)
        match totalcodes - allocatedcodes with
        | amount when amount < 0 -> Over(abs amount)
        | amount when amount > 0 -> Under(amount)
        | _ -> Exact 

    let trypartition (allocations:seq<PartitionEntry>) (vpp:VPP) =
        let allocsaslist = lazy (allocations |> List.ofSeq) //For algorithmic purposes
        let skiptake amount list' = // assumes called with nonnegative for now
            let rec f taken a l =
                match a with
                | 0 -> taken,l
                | _ -> (List.head l , List.tail l) |> fun (h,t) -> f (h::taken) (a-1) t
            f [] amount list'
        let tovpp (entry:PartitionEntry) codes = 
            {vpp with FileName=entry.Name; VPPCodes=codes; CodesRedeemed=0; CodesRemaining=entry.Allocation; CodesPurchased=entry.Allocation}
        let rec f codes (entries:PartitionEntry list) vpps  =
            match entries with
            | [] -> vpps
            | entry::remainingentries -> 
                skiptake entry.Allocation codes
                |> fun (allocated,remainingcodes) -> f remainingcodes remainingentries (tovpp entry allocated::vpps) 
        match measure' vpp allocations with
        | Over amount ->
            do onpartitionoverallocation.Trigger(null,new AllocationErrorEventArgs(amount))
            None
        | Under amount -> 
            do onpartitionunderallocation.Trigger(null,new AllocationErrorEventArgs(amount))
            None
        | Exact -> Some (f vpp.AvailableVPPCodes (allocsaslist.Force()) [])

    //Events
    [<CLIEvent>] member this.OnAddDuplicateName = onaddduplicatename.Publish
    [<CLIEvent>] member this.OnAddInvalidFileName = onaddinvalidfilename.Publish
    [<CLIEvent>] member this.OnAddOverAllocation = onaddoverallocation.Publish
    [<CLIEvent>] member this.OnPartitionUnderAllocation = onpartitionunderallocation.Publish
    [<CLIEvent>] member this.OnPartitionOverAllocation = onpartitionoverallocation.Publish

    member this.VPP = VPP

    member this.Add(newentry) =
        let isvalidname (entry:PartitionEntry) = 
            match entry.Name.ToCharArray()|> Set.ofArray |> Set.intersect invalidfilechars |> Set.isEmpty with
            | true -> Some entry
            | false -> 
                do onaddinvalidfilename.Trigger(this,new PartitionErrorEventArgs(entry))
                None
        let isduplicate (entry:PartitionEntry) = 
            match partitions |> Seq.exists(fun curr -> curr.Name = entry.Name) with
            | true -> 
                do onaddduplicatename.Trigger(this,new PartitionErrorEventArgs(entry))
                None
            | false -> Some entry
        let allocmeasure (entry:PartitionEntry) =
            let wantedpartitions = Seq.singleton entry |> Seq.append partitions
            match measure' this.VPP wantedpartitions with
            | Over amount -> do onaddoverallocation.Trigger(this,new EntryAllocationErrorEventArgs(amount,entry))
            | _ -> partitions.Add(newentry)
        do newentry |> isvalidname |> Option.bind(isduplicate) |> Option.iter(allocmeasure)

    member this.WriteToXLS(directorypath) =
        match trypartition partitions this.VPP with 
        | Some(vpps) ->  vpps |> Seq.iter(fun vpp -> savevpptoxls vpp (Path.Combine(directorypath,vpp.FileName + ".xls")))
        | None -> () //Need to be refactored out
    member this.WriteToXLS(directorychooser:DirectoryChooser) =
        match trypartition partitions this.VPP with 
        | Some(vpps) -> 
            let directorypath = directorychooser.Invoke()
            vpps |> Seq.iter(fun vpp -> savevpptoxls vpp (Path.Combine(directorypath,vpp.FileName + ".xls")))
        | None -> () //Need to be refactored out