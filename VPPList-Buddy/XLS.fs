module VPPListBuddy.XLS

open System
open System.IO
open VPPListBuddy.VPP
open VPPListBuddy.IO

let xlsreadcell (sheet : Worksheet) (index : int * int) =
    index |> fun (h,v) -> sheet.Cells.Item(h,v)

let xlstrystring (sheet : Worksheet) (index : int * int) =
    match (xlsreadcell sheet index).StringValue with
    | text when String.IsNullOrWhiteSpace(text) -> None
    | text -> Some(text)

let xlstryint ws = 
    xlstrystring ws >> function
    | Some(str) -> 
        match str |> Int32.TryParse with
        | (true,i) -> Some(i)
        | (false,_) -> None
    | None -> None

let xlswritecell (sheet : Worksheet) (index : int * int) (obj:Object)  = 
    do index |> fun (h,v) -> sheet.Cells.Item(h,v) <- new Cell(obj) //How safe is this?

let xlswritetext (sheet : Worksheet) (index : int * int) (text:string) =
    xlswritecell sheet index text

let xlsworksheet name =
    new Worksheet(name)

let xlsopenfile (path : string) =
    try
        Some(path |> Workbook.Load)
    with
    | ex  -> None

let attachsheet sheet (wb : Workbook) =
    do wb.Worksheets.Add(sheet)
    wb

let xlssaveworkbook (path : string) (wb : Workbook)  = 
    wb.Save(path)
    wb

let xlssaveworksheet (path : string) (ws : Worksheet) =
    let wb = Workbook()
    do 
        wb.Worksheets.Add(ws)
        wb.Save(path)

let xlstestempty (xls: Workbook) = if xls.Worksheets.Count = 0 then None else Some(xls)
let xlsextractsheet (xls : Workbook) sheetindex =
    match xls.Worksheets.Count , sheetindex with
    | (0,_) -> None
    | (_,index) when 0 > index -> None
    | (sheetcount,index) when sheetcount <= index ->  None
    | _ -> Some(xls.Worksheets.Item(sheetindex))

let xlsextractsheet' sheetindex xls = xlsextractsheet xls sheetindex

let mapvppdatacell (spreadsheet : VPP) =
    let vppdata = [
            (0,0,"Volume Purchase Codes");(0,7,"ProductType");(0,8,spreadsheet.ProductType);(0,9,"|");(0,10,"AdamId:");(0,11,spreadsheet.AdamID);
            (1,0,"Order ID"); (1,2,spreadsheet.OrderID);
            (2,0,"Product"); (2,2,spreadsheet.Product);
            (3,0,"Purchaser"); (3,2,spreadsheet.Purchaser);
            (4,0,"Codes Purchased"); (4,2,spreadsheet.CodesPurchased |> string);
            (5,0,"Codes Redeemed"); (5,2,spreadsheet.CodesRedeemed |> string);
            (6,0,"Codes Remaining"); (6,2,spreadsheet.CodesRemaining |> string);
            (9,0,"Code");(9,2,"Code Redemption Link")]
    spreadsheet.VPPCodes //map cells to spreadsheet
    |> List.mapi(fun vindex code -> 
        let startindex = 10
        (vindex + startindex,0,code.Code),(vindex+startindex,2,code.URL))
    |> List.collect(fun (code,url) -> [code;url])
    |> List.append vppdata


let parsevppspreadsheet cellintreader (cellreader :  int * int -> string option) = //need to verify spreadsheet
    let tryint cell = match cellintreader cell with | Some(int') -> int' | None -> 0
    let trytext cell = match cellreader cell with | Some(text) -> text | None -> ""
    let extractinfo vpp =
        vpp 
        |> fun vpp -> {vpp with ProductType = trytext (0,8)}
        |> fun vpp -> {vpp with AdamID = trytext (0,11)}
        |> fun vpp -> {vpp with OrderID = trytext (1,2)}
        |> fun vpp -> {vpp with Product = trytext (2,2)}
        |> fun vpp -> {vpp with Purchaser = trytext (3,2)}
        |> fun vpp -> {vpp with CodesPurchased = tryint (4,2)}
        |> fun vpp -> {vpp with CodesRedeemed = tryint (5,2)}
        |> fun vpp -> {vpp with CodesRemaining = tryint (6,2)}
    let extractcodes =
        let rec extract start (vpp : VPP)  =
            let vppentry = cellreader (start,0) , cellreader (start,2)  
            match vppentry with 
            | Some(code) , Some(url) -> extract (start + 1) ({Code=code;URL=url} |> fun r -> {vpp with VPPCodes = r :: vpp.VPPCodes })
            | _ , _ -> vpp
        extract 10
    let verify (vpp:VPP) =
        let badtotal = vpp.VPPCodes.Length <> vpp.CodesPurchased 
        let badavailable = vpp.AvailableVPPCodes.Length <> vpp.CodesRemaining
        let badreedeemed = vpp.ReedeemedVPPCodes.Length <> vpp.CodesRedeemed
        let anybad = [badtotal;badavailable;badreedeemed] |> List.exists(id)
        if anybad then None else Some(vpp) 
    extractinfo blankspreadsheet |> extractcodes |> verify

let xlsparsevpp worksheet =
    let celltextreader = xlstrystring worksheet
    let cellintreader = xlstryint worksheet
    parsevppspreadsheet cellintreader celltextreader

let savevpptoxls vpp path =
    let xlsserializer = xlssaveworksheet path
    let write = 
        fun xlsserializer -> savespreadsheet (Worksheet("Sheet1")) xlswritetext xlsserializer mapvppdatacell
    write xlsserializer vpp