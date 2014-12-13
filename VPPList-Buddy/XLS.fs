module VPPListBuddy.XLS

open System
open System.IO
open VPPListBuddy.VPP
open VPPListBuddy.IO
open NPOI.HSSF
open NPOI.HSSF.UserModel
open NPOI.SS.UserModel

let xlsopenfile path =
    use fs = path |> File.OpenRead
    try
        Some(HSSFWorkbook(fs))
    with
    | ex  -> None

let xlstestempty (xls: IWorkbook) = if xls.NumberOfSheets = 0 then None else Some(xls)

let attachsheet sheet (wb : IWorkbook) =
    do wb.Add(sheet)
    wb

let xlsworksheet name = //Not specifically an xls worksheet, but doesn't matter for this application
    let wb = new HSSFWorkbook()
    let ws = wb.CreateSheet() //All sheets need to be attached to a workbook I suppose?
    do wb.SetSheetName(0,name)
    ws

let xlsreadcell (sheet : ISheet) (index : int * int) = 
    index |> fun (v,h) -> sheet.GetRow(v).GetCell(h)

let xlstrystring (sheet : ISheet) (index : int * int) =
    match (xlsreadcell sheet index) with
    | null -> None
    | cell when String.IsNullOrWhiteSpace(cell.ToString()) -> None //Method StringCellValue throws exception on numeric values 
    | cell -> Some(cell.ToString())

let xlstryint ws = 
    xlstrystring ws >> function
    | Some(str) -> 
        match str |> Int32.TryParse with
        | (true,i) -> Some(i)
        | (false,_) -> None
    | None -> None

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


let xlsextractsheet (xls : IWorkbook) sheetindex =
    match xls.NumberOfSheets , sheetindex with
    | (0,_) -> None
    | (_,index) when 0 > index -> None
    | (sheetcount,index) when sheetcount <= index ->  None
    | _ -> Some(xls.GetSheetAt(sheetindex))

let xlsextractsheet' sheetindex xls = xlsextractsheet xls sheetindex

let xlsparsevpp worksheet =
    let celltextreader = xlstrystring worksheet
    let cellintreader = xlstryint worksheet
    parsevppspreadsheet cellintreader celltextreader

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

(*let xlswritecell (sheet : HSSFSheet) (index : int * int) (obj:Object)  = 
    //do index |> fun (h,v) -> sheet.Cells.Item(h,v) <- new Cell(obj) //How safe is this?*)

let xlswritetext (sheet : ISheet) (index : int * int) (text:string) =
    do index |> fun (v,h) -> sheet.CreateRow(v).CreateCell(h).SetCellValue(text)


let xlssaveworksheet (path : string) (ws : ISheet) =
    use fs = File.Create(path)
    let wb = new HSSFWorkbook()
    do 
        wb.Add(ws)
        wb.Write(fs)

let xlssaveworkbook (path : string) (wb : HSSFWorkbook)  = 
    use fs = File.Create(path)
    wb.Write(fs)

let savevpptoxls vpp path =
    let xlsserializer = xlssaveworksheet path
    let write = 
        fun xlsserializer -> savespreadsheet (xlsworksheet "sheet1") xlswritetext xlsserializer mapvppdatacell
    write xlsserializer vpp
