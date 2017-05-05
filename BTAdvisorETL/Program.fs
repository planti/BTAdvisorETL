//
//  Program.fs
//
//  Author:
//       Ivan <ivan.maffioli@sdgitaly.it>
//
// 13/01/2017 - 16:07
//
//

open System
open System.IO
open System.Text.RegularExpressions
open Chessie.ErrorHandling // per adesso inglobato come sorgente
open Common
open BtAdvisor
open UvetAmex
open FSharp.Data
open FSharp.Data.SqlClient

/// Lancia la sp che elabora i dati della staging area per inserire nelle tabelle definitive
let postFromStaging (opt:CommandLineOprtions) =
    let post_ () =
        if opt.operazione = Completa then
            use cmd = new Btadvisor.dbo.LoadTransactionData(targetConnectionString)
            verboseOutput opt.verbose "Post dati da staging nelle tabelle definitive ..."
            // Result contiene il numero di record affected.
            // Nel caso la sp non torna nulla e inizia con un set nocount on
            // result = -1 
            // in questo caso ignoro il valore di result
            let result = cmd.AsyncExecute (Parte = opt.idparte, idSorgente = opt.idadv, periodo = opt.periodo, anno = opt.anno)
                         |> Async.RunSynchronously 
            
            verboseOutput opt.verbose "post terminato " 
        
    tryF post_ DbUpdateFailure

// --------------------------------------------------------------------------

/// Parsa un pattern di parametri fatto così
/// program -f Nomefile [-t csv | excel | UAExcel] [-a idadv] [-p idparte] [-v] [-post periodo anno]
let parseCommandLine args = 
    // Helper recursive parsing function
    let rec parse args optionsSoFar = 
        match args with
        | [] -> optionsSoFar

        | "-v" :: xs -> 
            let newOptionSoFar = {optionsSoFar with verbose = VerboseOutput}
            parse xs newOptionSoFar

        | "-a" :: xs ->
            match xs with 
            | id  :: xss ->
                let m = Regex.Match(id, "\d+") in
                    if m.Success then 
                        let newOptionsSoFar = {optionsSoFar with idadv = int(id)}
                        parse xss newOptionsSoFar
                    else 
                        eprintfn "Argomento Errato per idadv: %s default 1" id
                        parse xss optionsSoFar
            | _ -> 
                eprintfn "-a vuole un idadv, assume 1"
                parse xs optionsSoFar

        | "-p" :: xs -> 
            match xs with 
            | id  :: xss ->
                let m = Regex.Match(id, "\d+") in
                    if m.Success then 
                        let newOptionsSoFar = {optionsSoFar with idparte = int(id)}
                        parse xss newOptionsSoFar
                    else 
                        eprintfn "Argomento Errato per idparte: %s, default 1" id
                        parse xss optionsSoFar
            | _ -> 
                eprintfn "-p vuole un idparte, assume 1"
                parse xs optionsSoFar
        
        | "-f" :: xs ->
            match xs with
            | fileName :: xss ->
                let newOptionsSoFar = {optionsSoFar with file = fileName}
                parse xss newOptionsSoFar
            | _ ->
                eprintfn "-f vuole un nome di file"   
                parse xs optionsSoFar   
    
        | "-t" :: xs ->
            match xs with
            | "excel" :: xss ->
                let newOptionsSoFar = {optionsSoFar with fileFormat = Excel}
                parse xss newOptionsSoFar
            | "UAexcel" :: xss ->
                let newOptionsSoFar = {optionsSoFar with fileFormat = UAExcel}
                parse xss newOptionsSoFar
            | "csv" :: xss ->
                let newOptionsSoFar = {optionsSoFar with fileFormat = Csv}
                parse xss newOptionsSoFar
            | _ ->
                eprintfn "Formato non riconosciuto, si assume Csv"   
                parse xs optionsSoFar   

        | "-post" :: xs ->
            match xs with
            | periodo :: anno :: xss ->
                let p = Regex.Match(periodo, "[0-4]") in
                    let a = Regex.Match(anno, "\d+") in
                        if p.Success && a.Success then 
                            let newOptionsSoFar = {optionsSoFar with operazione = Completa; periodo = int(periodo); anno = int(anno)}
                            parse xss newOptionsSoFar
                        else 
                            eprintfn "Argomento Errato per periodo e anno: %s %s, default 0 2017" periodo anno
                            parse xss optionsSoFar
            | _ -> 
                eprintfn "-post vuole un periodo e un anno" 
                parse xs optionsSoFar  

        | x :: xs ->
            eprintfn "Opzione %s non è riconosciuta, ignorata" x
            parse xs optionsSoFar 

    let defaultOptions = {
        verbose = TerseOutput;
        file = "C:\Users\imaffioli.SDGITALY\Desktop\BTAdvisor\pippo.xls";
        fileFormat = UAExcel;
        idadv = 1;
        idparte = 1;
        operazione = Staging;
        periodo = 0;
        anno = 2017
    }

    if List.length args = 0 then 
        eprintfn "uso: BTAdvisorETL -f Nomefile [-t csv | excel | UAexcel | CExcel] [-a idadv] [-p idparte] [-v] [-post periodo anno]"
        eprintfn "utilizza le opzioni di default "
        defaultOptions
    else 
        parse args defaultOptions

// --------------------------------------------------------------------------

/// Controlla se il file da elaborare c'è
let checkFileExists opt =
    if not (File.Exists(opt.file)) then
        fail (GenericText (sprintf "Il file %s non esiste" opt.file))
    else opt |> ok

// --------------------------------------------------------------------------   

[<EntryPoint>]
let main argv =
    let options = argv |> List.ofArray |> parseCommandLine
    let result = match options.fileFormat with
                 | Csv     -> checkFileExists options >>= readCsv >>= postCsvData options >>= postFromStaging
                 | Excel   -> checkFileExists options >>= readExcel >>= postExcelData options >>= postFromStaging
                 | UAExcel -> checkFileExists options >>= readUAExcel >>= postUAExcelData options >>= postFromStaging
                 | CExcel  -> checkFileExists options >>=

    match result with
    | Ok _ -> verboseOutput options.verbose " Elaborazione terminata" ; 0
    | Bad errs -> eprintfn "%A" errs ; -1
