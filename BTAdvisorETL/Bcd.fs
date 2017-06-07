/// BCD
/// Modulo che contiene il codice per caricare nella staging area un foglio excel
/// nel formato utilizzato da BCD
///
module Bcd
    open System
    open System.Configuration
    open System.Text.RegularExpressions
    open FSharp.ExcelProvider
    open Common
    open ShellProgressBar 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open TryParser

    type BCDExcelDataSource = ExcelFile<"BCD.xlsx", HasHeaders=true, ForceString=true>

// --------------------------------------------------------------------------

    let mapTransacrtionType bgtRimborso tipoTariffa =
        match bgtRimborso with
        | "Y" -> Some "Refund"
        | "N" when (Regex.Match(tipoTariffa, "^Ticket Re-Issue")).Success -> Some "NewEmisson"
        | "N" 
        | _ -> Some "Emission"

// --------------------------------------------------------------------------

    /// Determina il tipo di routing
    let mapRoutingType originalType routing product =
        match product with 
        | "Aereo" 
        | "Low Cost" ->
            match originalType with
            | "Andata" when (Regex.Match(routing, "^\w{3}/\w{3}$")).Success -> Some "OneWay"
            | "Andata/Ritorno" when (Regex.Match(routing, "^\w{3}/\w{3}/\w{3}$")).Success -> Some "RoundTrip"
            | _ -> Some "MultiTratta"
        | "Ferroviario" 
        | "Ferroviario Elettronico" -> Some "OneWay"
        | _ -> None

    let adjustRailRouting originalRouting destination =
        // se il routing ha una barra sola
        if Regex.Match(originalRouting, "^[\w\s,.]+/[\w\s,.]+$").Success then
            match destination with
            | "-" -> Some originalRouting
            | _ -> Some (originalRouting.Substring(0,originalRouting.IndexOf('/')) + string '/' + destination)
        else Some originalRouting

// --------------------------------------------------------------------------

    /// Legge il file excel BCD
    let readBCDExcel opt =
        let read_ () = new BCDExcelDataSource(opt.file)
        verboseOutput opt.verbose "Lettura del file Excel BCD in corso ..."
        tryF read_ ExcelFileAccessFailure

// --------------------------------------------------------------------------
    
    /// Salva i dati nella tabella di staging operando opportune trasformazioni
    let postBCDExcelData opt (t:BCDExcelDataSource) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.Cliente)
    
            use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor) 
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sul DB",ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDateWithFormat row.Data_Registrazione "MM/dd/yyyy hh:mm:ss",
                        legalentityid   = Some row.Cod_Cliente,
                        legalentityname = Some row.Cliente, 
                        transactiontype = mapTransacrtionType row.Bgt_Rimborso row.Tipo_Tariffa,
                        product         = (match row.TpServizioDettaglio with
                                          | "Aereo" 
                                          | "Low Cost" -> Some "AIR"
                                          | "Autonoleggio" -> Some "CAR"
                                          | "Hotel" -> Some "HOT"
                                          | "Ferroviario Elettronico" -> Some "RAV"
                                          | "Ferroviario" -> Some "RAI"                                       
                                          | _ -> Some "VAR"),
                        documentno      = Some row.Num_Documento, 
                        tickettype      = (match row.TpServizioDettaglio with
                                          | "Aereo"                   
                                          | "Ferroviario Elettronico" -> Some "Eticket"
                                          | "Low Cost" -> Some "LowCost"
                                          | _ -> None), 
                        departuredate   = parseDateWithFormat row.Data_Partenza "MM/dd/yyyy hh:mm:ss", 
                        supplier        = (match row.TpServizioDettaglio with
                                          | "Ferroviario" | "Ferroviario Elettronico" -> Some row.GruppoVettori                                    
                                          | _  -> Some row.Fornitore),
                        airlinecode     = None, 
                        origin          = (match row.TpServizioDettaglio with
                                          | "Aereo"  
                                          | "Low Cost" -> Some row.Partenza_Cod
                                          | "Hotel"    -> Some row.Destinazione
                                          | _          -> Some row.Citta_Partenza),
                        origincountrycd = None,                       
                        destination     = (match row.TpServizioDettaglio with
                                          | "Aereo"   
                                          | "Low Cost" -> Some row.Destinazione_Cod
                                          | "Hotel"    -> None
                                          | _          -> Some row.Destinazione),                        
                        destcountrycd   = (match row.TpServizioDettaglio with
                                          | "Hotel" -> None
                                          | _       -> Some row.Nazione),
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = (match row.TpServizioDettaglio with
                                          | "Aereo"    
                                          | "Low Cost" -> Some row.Itinerario
                                          | "Ferroviario" | "Ferroviario Elettronico" -> adjustRailRouting row.Itinerario row.Destinazione
                                          | _ -> None), 
                        classofservices = (match row.TpServizioDettaglio with
                                          | "Aereo"    
                                          | "Low Cost" 
                                          | "Ferroviario" 
                                          | "Ferroviario Elettronico" -> Some row.DescrizioneClasse
                                          | _ -> None), 
                        roomtype        = None,
                        roomnight       = (match row.TpServizioDettaglio with 
                                          | "Hotel" -> parseInt row.``Num Notti``
                                          | _ -> None), 
                        daysofrent      = (match row.TpServizioDettaglio with 
                                          | "Autonoleggio" -> parseInt row.``Num Notti``
                                          | _ -> None), 
                        triptype        = (match row.Itinerario_Tipo with
                                          | "Nazionale" -> Some "Domestic"
                                          | "Continentale" -> 
                                                match row.Zona with
                                                | "EUROPA"
                                                | "ITALIA" -> Some "European"
                                                | _ -> Some "International"
                                          | "Intercontinentale" -> Some "International"
                                          | _ ->
                                          // In questo caso per attribuire correttmente il valore occorrerebbe 
                                          // considerare paese per paese
                                                match row.Nazione with
                                                | "ITALY" -> Some "Domestic"
                                                | _ -> Some "International"), 
                                        // Dopo telefonata con Dario
                        fullfare        = (match (parseDecimalX row.``Importo Commissionabile`` , parseDecimalX row.Saving) with
                                          | (Some x , Some y) -> Some (x+y)
                                          | (Some x , None) -> Some x
                                          | (None , Some y) -> Some 0.0M
                                          | _ -> Some 0.0M), 
                                        // Non c'è lo fisso all'importo pagato
                        lowfare         = (match parseDecimalX row.``Importo Commissionabile`` with
                                          | Some x -> Some x
                                          | None -> Some 0.0M),  
                        farepaid        = (match parseDecimalX row.``Importo Commissionabile`` with
                                          | Some x -> Some x
                                          | None -> Some 0.0M),  
                        reference       = None, 
                        farebasis       = None, 
                        fee             = Some 0.0M, 
                        tax             = parseDecimalX row.``Tasse Aeree``, 
                        routingtype     = mapRoutingType row.Bgt_Andata_Ritorno row.Itinerario row.TpServizioDettaglio, 
                        mileage         = None,    
                        marketcountry   = Some "IT",
                        pax             = Some row.Nome_Passeggero,
                        grade           = None, 
                        cdc             = Some row.Centro_Costo, 
                        aux             = None, 
                        bookingno       = None, 
                        invoiceno       = None, 
                        inpolicy        = None, 
                        reason          = Some row.Motivazione, 
                        requestdate     = parseDateWithFormat row.Data_Registrazione "MM/dd/yyyy hh:mm:ss", 
                        channnel        = None, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura dati sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure   
