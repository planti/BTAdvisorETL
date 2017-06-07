/// Cisalpina
/// Modulo che contiene il codice per caricare nella staging area un foglio excel
/// nel formato utilizzato da kpmg
///

/// Cisalpina
/// Modulo che contiene il codice per caricare nella staging area un foglio excel
/// nel formato utilizzato da kpmg
///
module Cisalpina.newExcel

    open System
    open System.Configuration
    open System.Text.RegularExpressions
    open FSharp.ExcelProvider
    open Common
    open ShellProgressBar 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open TryParser
    open Microsoft.Office.Interop

    /// sorgente dati Cisalpina
    /// I file di questa sorgente sono in diversi fogli excel dentro un unico workbook.
    /// Ogni foglio corrisponde ad un tipo di prodotto.
    /// I file excel possono contenere anche solo parte dei prodotti

    /// Tipo per accedere al foglio contenente il prodotto Air, Sheet: Air
    type CisalpinaAirDS = ExcelFile<"Cisalpina_Passenger.xls", "Air", Range="C10", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Hotel, Sheet: Air
    type CisalpinaHotDS = ExcelFile<"Cisalpina_Passenger.xls", "Hotel", Range="C10", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Car, Sheet: Air
    type CisalpinaCarDS = ExcelFile<"Cisalpina_Passenger.xls", "Car", Range="C10", HasHeaders=true, ForceString=true>

    /// Tipo per accedere al foglio contenente il prodotto Rail, Sheet: Air
    type CisalpinaRaiDS = ExcelFile<"Cisalpina_Passenger.xls", "Railway", Range="C10", HasHeaders=true, ForceString=true>

    /// Un worbok è composto da vari fogli ogni foglio ha un suo data provider
    /// questo tipo rappresenta l'insieme dei data provider
    type WorkBook = {
        Air: CisalpinaAirDS;
        Hot: CisalpinaHotDS;
        Car: CisalpinaCarDS;
        Rai: CisalpinaRaiDS;
    }

// --------------------------------------------------------------------------

    /// legge tutti i fogli del workbook
    let readCExcelSheets opt =
        verboseOutput opt.verbose "Lettura del file Excel cisalpina in corso ..."
        let readAll () = 
            let workbook = { Air = null; Car = null; Hot = null; Rai = null }
            // Legge gli sheet
            let rec read wb sl =
                match sl with
                | []              -> wb
                | "Air" :: xs     -> 
                    verboseOutput opt.verbose "Lettura dello sheet Air ..." 
                    read {wb with Air = new CisalpinaAirDS(opt.file)} xs
                | "Hotel" :: xs   -> 
                    verboseOutput opt.verbose "Lettura dello sheet Hotel ..." 
                    read {wb with Hot = new CisalpinaHotDS(opt.file)} xs
                | "Car" :: xs     -> 
                    verboseOutput opt.verbose "Lettura dello sheet Car ..." 
                    read {wb with Car = new CisalpinaCarDS(opt.file)} xs
                | "Railway" :: xs -> 
                    verboseOutput opt.verbose "Lettura dello sheet Railway ..." 
                    read {wb with Rai = new CisalpinaRaiDS(opt.file)} xs
                | _ :: xs         -> read wb xs
            read workbook (getSheets opt)
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
        | "One Way" when (Regex.Match(routing, "^\w{3}/\w{3}$")).Success -> "OneWay"
        | "One Way" -> "MultiTratta"
        | "Round Trip" when (Regex.Match(routing, "^\w{3}/\w{3}/\w{3}$")).Success -> "RoundTrip"
        | "Round Trip" -> "MultiTratta"
        | _ -> "?" // In questo modo viene scartato

// --------------------------------------------------------------------------

    /// Salva in staging la parte AIR
    let postCisalpinaExcelAirData opt (t:CisalpinaAirDS) =
        // salta eventuali righe vuote e dei totali
        let rows = t.Data |> Seq.where (fun x -> not <| isNull x.``Issuing Date``)
   
        use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor) 
        let count = rows |> Seq.length
        use pbar = new ProgressBar(count, "Scrittura sheet Air sul DB",ConsoleColor.Yellow)
        for row in rows do
            let recordsInserted =
                cmd.Execute(
                    idadv           = Some opt.idadv,
                    idparte         = Some opt.idparte,  
                    issuedate       = parseExcelDate row.``Issuing Date`` MustBeADate, 
                    legalentityid   = Some row.Company,
                    legalentityname = Some row.Company, 
                    // questa assegnazione è temporanea
                    transactiontype = (match parseDecimalX row.Amount with
                                      | Some x when (x > 0.0M && row.``Booking Type`` =  "RET") -> Some "Reemission" 
                                      | Some x when (x >= 0.0M && row.``Booking Type`` <> "RET") -> Some "Emission" 
                                      | Some x -> Some "Refund"
                                      | None -> Some "?"), // In questo modo viene scartato
                    product         = Some "AIR",  
                    documentno      = Some row.``Tkt Number``, 
                    tickettype      = Some (mapTicketType row.``Tkt Number``), 
                    departuredate   = parseExcelDate row.``Dept Date`` MustBeADate, 
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
                    requestdate     = parseExcelDate row.``Request Date`` CanBeNull, 
                    channnel        = None, 
                    loaddate        = Some DateTime.Today)
            pbar.Tick("Scrittura sheet Air sul DB ... ")


// --------------------------------------------------------------------------

    /// Salva in staging la parte HOT
    let postCisalpinaExcelHotData opt (t:CisalpinaHotDS) =
        // salta eventuali righe vuote e dei totali
        let rows = t.Data |> Seq.where (fun x -> (not <| isNull x.``Issuing Date``) && 
                                                 (not <| isNull x.Hotel) &&
                                                 (not <| isNull x.In))
    
        use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor)
        let count = rows |> Seq.length
        use pbar = new ProgressBar(count, "Scrittura sheet Hotel sul DB" ,ConsoleColor.Yellow)
        for row in rows do
            let recordsInserted =
                cmd.Execute(
                    idadv           = Some opt.idadv,
                    idparte         = Some opt.idparte,  
                    issuedate       = parseExcelDate row.``Issuing Date`` MustBeADate, 
                    legalentityid   = Some row.Company,
                    legalentityname = Some row.Company, 
                    transactiontype = (match parseDecimalX row.Amount with
                                      | Some x when x > 0.0M -> Some "Emission" 
                                      | Some x               -> Some "Refund"
                                      | None -> Some "?"), // In questo modo viene scartato
                    product         = Some "HOT",
                    documentno      = Some row.``Voucher Nr.``, 
                    tickettype      = None, 
                    departuredate   = parseExcelDate row.In MustBeADate, 
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
                    requestdate     = parseExcelDate row.``Request Date`` CanBeNull, 
                    channnel        = None, 
                    loaddate        = Some DateTime.Today)
            pbar.Tick("Scrittura sheet Hotel sul DB in corso ... ")

// --------------------------------------------------------------------------

    /// Salva in staging la parte RAI
    let postCisalpinaExcelRaiData opt (t:CisalpinaRaiDS) =
        // salta eventuali righe vuote e dei totali
        let rows = t.Data |> Seq.where (fun x -> not <| isNull x.``Issuing Date``)
    
        use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor)
        let count = rows |> Seq.length
        use pbar = new ProgressBar(count, "Scrittura sheet Railway sul DB", ConsoleColor.Yellow)
        for row in rows do
            let recordsInserted =
                cmd.Execute(
                    idadv           = Some opt.idadv,
                    idparte         = Some opt.idparte,  
                    issuedate       = parseExcelDate row.``Issuing Date`` MustBeADate,  
                    legalentityid   = Some row.Company,
                    legalentityname = Some row.Company, 
                    transactiontype = (match parseDecimalX row.Amount with
                                      | Some x when x > 0.0M -> Some "Emission" 
                                      | Some x               -> Some "Refund"
                                      | None -> Some "?"), // In questo modo viene scartato
                    product         = Some "RAI",
                    documentno      = Some row.``Tkt Number``, 
                    tickettype      = Some "Eticket", 
                    departuredate   = parseExcelDate row.``Dept Date`` MustBeADate,  
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
                    requestdate     = parseExcelDate row.``Request Date`` CanBeNull, 
                    channnel        = None, 
                    loaddate        = Some DateTime.Today)
            pbar.Tick("Scrittura sheet Railway sul DB in corso ... ")

// --------------------------------------------------------------------------

    /// Salva in staging la parte CAR
    let postCisalpinaExcelCarData opt (t:CisalpinaCarDS) =
        // salta eventuali righe vuote e dei totali
        let rows = t.Data |> Seq.where (fun x -> not <| isNull x.``Issuing Date``)
    
        use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor)
        let count = rows |> Seq.length
        use pbar = new ProgressBar(count, "Scrittura sheet Car sul DB",ConsoleColor.Yellow)
        for row in rows do
            let recordsInserted =
                cmd.Execute(
                    idadv           = Some opt.idadv,
                    idparte         = Some opt.idparte,  
                    issuedate       = parseExcelDate row.``Issuing Date`` MustBeADate,  
                    legalentityid   = Some row.Company,
                    legalentityname = Some row.Company, 
                    transactiontype = (match parseDecimalX row.Amount with
                                      | Some x when x > 0.0M -> Some "Emission" 
                                      | Some x               -> Some "Refund"
                                      | None -> Some "?"), // In questo modo viene scartato
                    product         = Some "CAR",
                    documentno      = Some row.``Voucher Nr``, 
                    tickettype      = None, 
                    departuredate   = parseExcelDate row.``Pick Up Date`` MustBeADate,
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
                    requestdate     = parseExcelDate row.``Request Data`` CanBeNull, 
                    channnel        = None, 
                    loaddate        = Some DateTime.Today)
            pbar.Tick("Scrittura sheet Car sul DB in corso ... ")

 // --------------------------------------------------------------------------
       
    ///Salva tutti i worksheet    
    let postCExcel opt wb =
        // incapsula in una funzione senza parametri
        let post () = 
            // post vero e proprio in base ai fogli del workbook
            let rec post_ sl = 
                match sl with
                | []              -> opt
                | "Air" :: xs     -> postCisalpinaExcelAirData opt wb.Air; post_ xs
                | "Hotel" :: xs   -> postCisalpinaExcelHotData opt wb.Hot; post_ xs
                | "Car" :: xs     -> postCisalpinaExcelCarData opt wb.Car; post_ xs
                | "Railway" :: xs -> postCisalpinaExcelRaiData opt wb.Rai; post_ xs
                | _ :: xs         -> post_ xs
            post_ (getSheets opt)  // si potrebbe salvare il risultato intermedio di questo per non farlo due volte
        tryF post  DbUpdateFailure    
       