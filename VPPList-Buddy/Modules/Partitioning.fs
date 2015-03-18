open System
open VPPListBuddy.VPP
open VPPListBuddy.XLS

type AllocationMeasurement =
    | Over of Amount:int
    | Under of Amount:int
    | Exact

type AllocationsResult =
    | Ongoing of Stuff : VPP list * Rest : VPPCode list

let measure' amountof partitions vpp =
    let { CodesRemaining = totalcodes } = vpp
    let allocatedcodes = partitions |> Seq.sumBy amountof
    match totalcodes - allocatedcodes with
    | amount when amount < 0 -> Over(abs amount)
    | amount when amount > 0 -> Under(amount)
    | _ -> Exact

let partitionnewcodes filename codes originalvpp =
    let codeamount = List.length codes 
    {originalvpp with FileName=filename; VPPCodes=codes; CodesRedeemed=0; CodesRemaining=codeamount; CodesPurchased=codeamount}

let trypartition nameof amountof partitions vpp =
    let skiptake amount list' = // assumes called with nonnegative for now
        let rec f taken a l =
            match a with
            | 0 -> taken,l
            | _ -> (List.head l , List.tail l) |> fun (h,t) -> f (h::taken) (a-1) t
        f [] amount list'
    let rec f codes entries vpps =
        match entries with
        | [] -> vpps
        | entry::remainingentries -> 
            skiptake (amountof entry) codes
            |> fun (allocated,remainingcodes) -> f remainingcodes remainingentries (tovpp entry allocated::vpps)
    match measure' (fun a -> 2) partitions vpp with
    | Over amount as m -> m, None
        //wrap in somehting that takes the return and then werror
        //do onpartitionoverallocation.Trigger(null,new AllocationErrorEventArgs(amount))
    | Under amount as m-> m, None
        //do onpartitionunderallocation.Trigger(null,new AllocationErrorEventArgs(amount))
    | Exact as m -> m, Some (f vpp.AvailableVPPCodes (allocsaslist.Force()) [])
  