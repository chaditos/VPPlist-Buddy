module VPPListBuddy.VPP

open System

let (|Reedemed|Available|) urlentry = //The URL field contains info regarding VPP status
    match urlentry with
    | "redeemed" -> Reedemed
    | _ -> Available

type VPPCode = {Code : string ; URL : string}

type VPP = 
    {OrderID : string; Product : string; Purchaser : string;
    CodesPurchased : int; CodesRedeemed : int; CodesRemaining : int;
    ProductType : string; AdamID : string; VPPCodes : VPPCode list;
    FileName : string}
    member this.AvailableVPPCodes = this.VPPCodes |> List.filter(fun entry -> match entry.URL with | Available -> true | Reedemed -> false)
    member this.ReedeemedVPPCodes = this.VPPCodes |> List.filter(fun entry -> match entry.URL with | Available -> false | Reedemed -> true)


let blankspreadsheet = {OrderID=""; Product=""; Purchaser=""; CodesPurchased=0;CodesRedeemed=0; CodesRemaining=0; ProductType=""; AdamID="" ;VPPCodes= [];FileName=""}