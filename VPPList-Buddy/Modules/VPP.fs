module VPPListBuddy.VPP

open System


type VPPCode = {Code : string ; URL : string}

let (|Redeemed|Available|) vppcode = //The URL field contains info regarding VPP status
    let {URL=url} = vppcode
    match url with
    | "redeemed" -> Redeemed
    | _ -> Available

type VPP = //Calculate instead of just storing the value, makes it harder to write functions.
    {OrderID : string; Product : string; Purchaser : string;
    ProductType : string; AdamID : string; VPPCodes : VPPCode list;
    FileName : string}
    member this.CodesPurchased = List.length this.VPPCodes
    member this.CodesRedeemed = List.length <| (this.VPPCodes |> List.filter(function Redeemed -> true | Available -> false))
    member this.CodesRemaining = List.length <| (this.VPPCodes |> List.filter(function Available -> true | Redeemed -> false))
  
let blankspreadsheet = 
    {OrderID=""; Product=""; Purchaser="";
    ProductType=""; AdamID="" ;VPPCodes= [];FileName=""}

let vppavailablecodes (vpp:VPP) = vpp.VPPCodes |> List.filter(function Available -> true | Redeemed -> false)
let vppredeemedcodes (vpp:VPP) = vpp.VPPCodes |> List.filter(function Available -> false | Redeemed -> true)

let markfirstredeemed vpp = //Apple Configurator only works with one atleast one file redeemed.
    let {VPPCodes=codes; CodesRemaining=remaining} = vpp
    match codes with
    | [] -> vpp
    | h :: t ->
        match h with
        | Redeemed -> vpp
        | Available ->
            let marked = {h with URL="redeemed"}
            {vpp with VPPCodes = (marked :: t); CodesRedeemed=1;CodesRemaining=(remaining - 1) }
