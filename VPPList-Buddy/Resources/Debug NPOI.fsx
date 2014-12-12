#I "../bin/Debug/"
#r "NPOI"
#r "VPPList-Buddy.dll"

open System
open System.IO
open VPPListBuddy
open VPPListBuddy.VPP
open VPPListBuddy.IO
open VPPListBuddy.XLS
open NPOI.HSSF
open NPOI.HSSF.UserModel

let assetspath = __SOURCE_DIRECTORY__ 
let pathgoodvpp = Path.Combine(assetspath,"Sample.xls")
let pathbadtotals = Path.Combine(assetspath,"Incorrect Totals.xls")
let pathxlstemplate = Path.Combine(assetspath,"Template.xls")

let samplevpp = xlsopenfile pathgoodvpp

let debugcellreader cellreader index =
    samplevpp
    |> Option.iter( fun xls -> 
        let ws = xls.GetSheetAt(0)
        printfn "%A" (index |> cellreader ws))

let readcell = debugcellreader xlsreadcell
let readcellstring = debugcellreader xlstrystring
let readcellint = debugcellreader xlstryint

