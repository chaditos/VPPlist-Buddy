#I "../bin/Release/"
#r "NPOI.dll"
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
    |> Option.bind( fun xls -> 
        let ws = xls.GetSheetAt(0)
        cellreader ws index)//printfn "%A" (index |> cellreader ws))

//let readcell = debugcellreader xlsreadcell
let readcellstring = debugcellreader xlstrystring
let readcellint = debugcellreader xlstryint

let copyxls path =
    xlsopenfile path 
    |> Option.bind(xlstestempty)
    |> Option.iter(xlssaveworkbook (Path.Combine(assetspath,"Copy Test.xls")))

let testwritenofunc path =
    xlsopenfile path 
    |> Option.bind(fun wb -> wb.GetSheetAt(0).CreateRow(6).CreateCell(2).SetCellValue("HEYYYYY YALLL"); Some wb)//.GetRow(0).GetCell(0).SetCellValue("test") ; Some wb)
    |> Option.iter(xlssaveworkbook (Path.Combine(assetspath,"Copy Test.xls")))

let testsimplewrite() =
    let ws = xlsworksheet "test"
    do
        xlswritetext ws (0,0) "SSD" 
        xlssaveworksheet (Path.Combine(assetspath,"NPOI.xls")) ws
