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

//.Net Friendly Events
type AllocationErrorEventArgs(Amount:int) =
    inherit EventArgs()
    member this.Amount = Amount
type AllocationError = delegate of obj * AllocationErrorEventArgs -> unit

type PartitionErrorEventArgs(Entry:PartitionEntry) =
    inherit EventArgs()
    member this.PartitionEntry = Entry
type PartitionError = delegate of obj * PartitionErrorEventArgs -> unit

//Workflow 
type public PartitionWorkflow (VPP:VPP) = 
    let invaliddirchars = Set.ofArray (Path.GetInvalidPathChars())

    let invalidfilechars = Set.ofArray (Path.GetInvalidPathChars())

    let measure' (vpp:VPP) (partitions:seq<PartitionEntry>)  =
        let totalcodes = vpp.CodesRemaining
        let allocatedcodes = partitions |> Seq.sumBy(fun p -> p.Allocation)
        match totalcodes - allocatedcodes with
        | amount when amount < 0 -> Over(amount)
        | amount when amount > 0 -> Under(amount)
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
   
    let partitions = Collections.Generic.List<PartitionEntry>()

    //Events
    let onaddduplicatename = new Event<PartitionError,PartitionErrorEventArgs>()
    let onaddinvalidfilename = new Event<PartitionError,PartitionErrorEventArgs>()
    let onaddoverallocation = new Event<PartitionError,PartitionErrorEventArgs>()
    let onpartitionunderallocation = new Event<AllocationError,AllocationErrorEventArgs>()
    let onpartitionoverallocation = new Event<AllocationError,AllocationErrorEventArgs>()
    [<CLIEvent>] member this.OnAddDuplicateName = onaddduplicatename.Publish
    [<CLIEvent>] member this.OnAddInvalidFileName = onaddinvalidfilename.Publish
    [<CLIEvent>] member this.OnAddOverAllocation = onaddoverallocation.Publish
    [<CLIEvent>] member this.OnPartitionUnderAllocation = onpartitionunderallocation.Publish
    [<CLIEvent>] member this.OnPartitionOverAllocation = onpartitionunderallocation.Publish

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
            | Over amount -> onaddoverallocation.Trigger(this,new PartitionErrorEventArgs(entry))
            | _ -> partitions.Add(newentry)
        do newentry |> isvalidname |> Option.bind(isduplicate) |> Option.iter(allocmeasure)

    member this.WriteToXLS(directorypath) =
        match measure' this.VPP partitions with
        | Over amount -> onpartitionoverallocation.Trigger(this,new AllocationErrorEventArgs(amount))
        | Under amount -> onpartitionunderallocation.Trigger(this,new AllocationErrorEventArgs(amount))
        | Exact -> 
            match partition partitions this.VPP with 
            | Some(vpps) ->  vpps |> Seq.iter(fun vpp -> savevpptoxls vpp (Path.Combine(directorypath,vpp.FileName + ".xls")))
            | None -> () //Need to be refactored out