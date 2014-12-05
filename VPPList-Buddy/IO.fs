module VPPListBuddy.IO

open System
open System.IO
open VPPListBuddy.VPP

let fileexists path = if File.Exists(path) then Some(path) else None
let filehasextension extension path = if FileInfo(path).Extension = extension then Some(path) else None

let savespreadsheet serializer cellwriter filewriter vppcellmap (vppsheet:VPP) =
    let writecell = cellwriter serializer
    let writevpp spreadsheet =  vppcellmap spreadsheet |> List.iter(fun (h:int,v:int,text:string) -> writecell (h,v) text)
    do 
        writevpp vppsheet
        filewriter serializer //filewriter would already be checked for write capability and only need serializer