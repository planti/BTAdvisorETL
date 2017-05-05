/// Cisalpina
/// Modulo che contiene il codice per caricare nella staging area un foglio excel
/// nel formato utilizzato da kpmg
///
module Cisalpina

    open System
    open System.Text.RegularExpressions
    open FSharp.ExcelProvider
    open Common
    open ShellProgressBar 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open TryParser
    open Microsoft.Office.Interop

    /// sorgente dati Cisalpina
    /// I file di questa sorgente sono in divesi fogli excel dentro un unico workbook.
    /// Ogni foglio corrisponde ad un tipo di prodotto.
    /// I file excel possono contenere anche solo parte dei prodotti

    /// Tipo per accedere al foglio contenente il prodotto Air, Sheet: Air
    type CisalpinaAirDS = ExcelFile<"KPMG.xls", "Air", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Hotel, Sheet: Air
    type CisalpinaHotDS = ExcelFile<"KPMG.xls", "Hotel", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Car, Sheet: Air
    type CisalpinaCarDS = ExcelFile<"KPMG.xls", "Car", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Rail, Sheet: Air
    type CisalpinaRaiDS = ExcelFile<"KPMG.xls", "Railway", HasHeaders=true, ForceString=true>

    /// Un worbok
    type WorkBook = 
        | Air of CisalpinaAirDS
        | Hot of CisalpinaHotDS
        | Car of CisalpinaCarDS
        | Rai of CisalpinaRaiDS

// --------------------------------------------------------------------------

    // Il foglio excel con i dati potrebbe non contenere tutti i servizi.
    // La librera principale di accesso ai dati excel restituisce come default
    // il primo foglio se nella dichiarazione di tipo viene indicato un foglio inesistente.
    // Inoltre non consente l'enumerazione dei fogli presenti nel workbook perciò
    // si utilizza in questa piccola porzione, la libreria interop di Microsoft.

    /// Determina i fogli che ci sono
    let getSheets opt =
        let xlApp = new Excel.ApplicationClass() 
        let xlWorkBook = xlApp.Workbooks.Open(opt.file)
        let wse = xlWorkBook.Worksheets
        let wsl = [ for s in wse -> (s:?> Excel.Worksheet).Name ]
        xlWorkBook.Close()
        wsl

// --------------------------------------------------------------------------

    let readCExcelSheets opt =
        verboseOutput opt.verbose "Lettura del file Excel cisalpina in corso ..."
        let readAll () = 
            // Legge gli sheet
            let rec read dl sl =
                match sl with
                | []              -> dl
                | "Air" :: xs     -> read (Air (new CisalpinaAirDS(opt.file)) :: dl) xs
                | "Hotel" :: xs   -> read (Hot (new CisalpinaHotDS(opt.file)) :: dl) xs
                | "Car" :: xs     -> read (Car (new CisalpinaCarDS(opt.file)) :: dl) xs
                | "Railway" :: xs -> read (Rai (new CisalpinaRaiDS(opt.file)) :: dl) xs
                | _ :: rest       -> read dl rest
            read [] (getSheets opt)
        tryF readAll ExcelFileAccessFailure

// --------------------------------------------------------------------------

    /// Ricava il tipo di biglietto dalla stringa del campo Tkt Number
    let mapTicketType (ticketNo:String) =
        let comp = ticketNo.Split([|' '|]) |> List.ofArray
        match comp with
        | loc :: ticketType :: rest ->
            match ticketType with
            | "BSP" -> "Eticket"
            | _     -> "Lowcost"
        | _ -> "?"  // In questo modo viene scartato

// --------------------------------------------------------------------------

    /// Determina il tipo di routing
    let mapRoutingType routing originalType =
        match originalType with
        | "One Way" when (Regex.Match(routing, "\w{3}/\w{3}")).Success -> "OneWay"
        | "One Way" -> "MultiTratta"
        | "Round Trip" when (Regex.Match(routing, "\w{3}/\w{3}/\w{3}")).Success -> "RoundTrip"
        | "Round Trip" -> "MultiTratta"
        | _ -> "?" // In questo modo viene scartato

// --------------------------------------------------------------------------

    /// Salva in staging la parte AIR
    let postCisalpinaExcelAirData opt (t:CisalpinaAirDS) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.Company)
    
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sheet Air sul DB",ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.``Issuing Date``,
                        legalentityid   = Some row.Company,
                        legalentityname = Some row.Company, 
                        // questa assegnazione è temporanea
                        transactiontype = Some "Emission",
                        product         = Some "AIR",  
                        documentno      = Some row.``Tkt Number``, 
                        tickettype      = Some (mapTicketType row.``Tkt Number``), 
                        departuredate   = parseDate row.``Dept Date``, 
                        supplier        = Some row.Airline,
                        airlinecode     = None, 
                        origin          = Some row.Departure,
                        origincountrycd = Some row.``Dept Nation``,
                        destination     = Some row.Destination,
                        destcountrycd   = Some row.``Dest Nation``,
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = Some row.Routing, 
                        classofservices = Some row.Class, 
                        roomtype        = None, 
                        roomnight       = None,
                        daysofrent      = None, 
                        triptype        = (match row.Area with
                                          | "Nazionale"         -> Some "Domestic"
                                          | "Internazionale"    -> Some "International"
                                          | "Intercontinentale" -> Some "International"
                                          | _                   -> Some "?"), // In questo modo viene scartato
                        fullfare        = parseDecimalX row.``Reference Fare``, 
                        lowfare         = parseDecimalX row.``Proposed Fare``, 
                        farepaid        = parseDecimalX row.Amount, 
                        reference       = None, 
                        farebasis       = Some row.Class_Tkt, 
                        fee             = Some 0.0M, 
                        tax             = parseDecimalX row.Taxes, 
                        routingtype     = Some (mapRoutingType row.Routing row.``OW - RT``), 
                        mileage         = None,    
                        marketcountry   = Some "IT",
                        pax             = Some row.``Passenger Name``,
                        grade           = Some row.``Optional Field 2``, 
                        cdc             = Some row.``Cost Centre``, 
                        aux             = Some row.CID, 
                        bookingno       = None, 
                        invoiceno       = Some row.``Nr Bolla``, 
                        inpolicy        = None, 
                        reason          = Some row.``Reason Code``, 
                        requestdate     = parseDate row.``Request Date``,  
                        channnel        = None, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura sheet Air sul DB ... ")
            opt
        tryF post_ DbUpdateFailure   

// --------------------------------------------------------------------------

    /// Salva in staging la parte HOT
    let postCisalpinaExcelHotData opt (t:CisalpinaHotDS) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.Company)
    
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sheet Hotel sul DB" ,ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.``Issuing Date``,
                        legalentityid   = Some row.Company,
                        legalentityname = Some row.Company, 
                        transactiontype = (match parseDecimalX row.Amount with
                                          | Some x when x > 0.0M -> Some "Emission" 
                                          | Some x               -> Some "Refund"
                                          | None -> Some "?"), // In questo modo viene scartato
                        product         = Some "HOT",
                        documentno      = Some row.``Voucher Nr.``, 
                        tickettype      = None, 
                        departuredate   = parseDate row.In, 
                        supplier        = Some row.Hotel,
                        airlinecode     = None, 
                        origin          = Some row.City,
                        origincountrycd = Some row.Country,  
                        destination     = None,
                        destcountrycd   = None,
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = None, 
                        classofservices = Some row.Category,
                        roomtype        = Some row.Room, 
                        roomnight       = parseInt row.Nights,
                        daysofrent      = None, 
                        triptype        = (match row.Country with
                                          | "ITALIA"         -> Some "Domestic"
                                          | _                -> Some "International"), 
                        fullfare        = parseDecimalX row.Amount, 
                        lowfare         = parseDecimalX row.Amount, 
                        farepaid        = parseDecimalX row.Amount, 
                        reference       = None, 
                        farebasis       = None, 
                        fee             = Some 0.0M, 
                        tax             = Some 0.0M, 
                        routingtype     = None, 
                        mileage         = None,    
                        marketcountry   = Some "IT",
                        pax             = Some row.``Passenger Name``,
                        grade           = Some row.``Optional Field 2``, 
                        cdc             = Some row.``Cost Centre``, 
                        aux             = Some row.CID, 
                        bookingno       = None, 
                        invoiceno       = Some row.``Nr Bolla``, 
                        inpolicy        = None, 
                        reason          = Some row.``Reason Code``, 
                        requestdate     = parseDate row.``Request Date``,  
                        channnel        = None, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura sheet Hotel sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure   

// --------------------------------------------------------------------------

    /// Salva in staging la parte RAI
    let postCisalpinaExcelRaiData opt (t:CisalpinaRaiDS) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.Company)
    
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sheet Railway sul DB", ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.``Issuing Date``,
                        legalentityid   = Some row.Company,
                        legalentityname = Some row.Company, 
                        transactiontype = (match parseDecimalX row.Amount with
                                          | Some x when x > 0.0M -> Some "Emission" 
                                          | Some x               -> Some "Refund"
                                          | None -> Some "?"), // In questo modo viene scartato
                        product         = Some "RAI",
                        documentno      = Some row.``Tkt Number``, 
                        tickettype      = Some "Eticket", 
                        departuredate   = parseDate row.``Dept Date``, 
                        supplier        = Some row.Railway,
                        airlinecode     = None, 
                        origin          = Some row.Departure,
                        origincountrycd = None,  
                        destination     = Some row.Destination,
                        destcountrycd   = None,
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = Some row.Routing, 
                        classofservices = Some row.Class,
                        roomtype        = None, 
                        roomnight       = None,
                        daysofrent      = None, 
                        triptype        = (match row.Area with
                                          | "Nazionale" -> Some "Domestic"
                                          | _           -> Some "International"), 
                        fullfare        = parseDecimalX row.Amount, 
                        lowfare         = parseDecimalX row.Amount, 
                        farepaid        = parseDecimalX row.Amount, 
                        reference       = None, 
                        farebasis       = None, 
                        fee             = Some 0.0M, 
                        tax             = Some 0.0M, 
                        routingtype     = None, 
                        mileage         = None,    
                        marketcountry   = Some "IT",
                        pax             = Some row.``Passenger Name``,
                        grade           = Some row.``Optional Field 2``, 
                        cdc             = Some row.``Cost Centre``, 
                        aux             = Some row.CID, 
                        bookingno       = None, 
                        invoiceno       = Some row.``Nr Bolla``, 
                        inpolicy        = None, 
                        reason          = None, 
                        requestdate     = parseDate row.``Request Date``,  
                        channnel        = None, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura sheet Railway sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure   

// --------------------------------------------------------------------------

    /// Salva in staging la parte CAR
    let postCisalpinaExcelCarData opt (t:CisalpinaCarDS) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.Company)
    
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sheet Car sul DB",ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.``Issuing Date``,
                        legalentityid   = Some row.Company,
                        legalentityname = Some row.Company, 
                        transactiontype = (match parseDecimalX row.Amount with
                                          | Some x when x > 0.0M -> Some "Emission" 
                                          | Some x               -> Some "Refund"
                                          | None -> Some "?"), // In questo modo viene scartato
                        product         = Some "CAR",
                        documentno      = Some row.``Voucher Nr``, 
                        tickettype      = None, 
                        departuredate   = parseDate row.``Pick Up Date``, 
                        supplier        = Some row.Rental,
                        airlinecode     = None, 
                        origin          = Some row.``Pick Up``,
                        origincountrycd = None,  
                        destination     = Some row.``Drop Off``,
                        destcountrycd   = None,
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = None, 
                        classofservices = None,
                        roomtype        = None, 
                        roomnight       = None,
                        daysofrent      = parseInt row.Days, 
                        triptype        = None, 
                        fullfare        = parseDecimalX row.Amount, 
                        lowfare         = parseDecimalX row.Amount, 
                        farepaid        = parseDecimalX row.Amount, 
                        reference       = None, 
                        farebasis       = None, 
                        fee             = Some 0.0M, 
                        tax             = Some 0.0M, 
                        routingtype     = None, 
                        mileage         = None,    
                        marketcountry   = Some "IT",
                        pax             = Some row.``Passenger Name``,
                        grade           = Some row.``Optional Field 2``, 
                        cdc             = Some row.``Cost Centre``, 
                        aux             = Some row.CID, 
                        bookingno       = None, 
                        invoiceno       = Some row.``Nr Bolla``, 
                        inpolicy        = None, 
                        reason          = None, 
                        requestdate     = parseDate row.``Request Data``,  
                        channnel        = None, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura sheet Car sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure 
        
        
    let postAll dl =
        "da finire"  
