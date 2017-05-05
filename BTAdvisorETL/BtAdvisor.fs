/// BtAdvisor
/// Modulo che contiene il codice per caricare nella staging area un file csv oppure un
/// foglio excel nel formato definito da BtAdvisor
///
module BtAdvisor
    open System
    open FSharp.ExcelProvider
    open Common
    open ShellProgressBar 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open TryParser

    /// Provider della sorgente dati csv standard
    type DataSource = 
        CsvProvider<
            Sample = "Sample.csv",
            Separators = ";", 
            AssumeMissingValues = true, 
            Culture = "it-IT", 
            PreferOptionals = true,
            // schema per le colonne non carattere
            Schema = "IssueDate=date, \
            DepartureDate=date, \
            RoomNight=int option, \
            DaysOfRent=int option, \
            FullFare=decimal, \
            LowFare=decimal, \
            FarePaid=decimal, \
            FEE=decimal, \
            Tax=decimal, \
            Mileage=float option, \
            RequestDate=date option">
    
    // Si assume excel con tutti i campi come testo
    // NOTA:
    // il type provider restituisce il separatore decimale come . 
    
    /// sorgente dati excel standard
    type ExcelDataSource = ExcelFile<"Sample.xls", HasHeaders=true, ForceString=true>

    /// Legge il file CSV nel formato BtAdvisor
    let readCsv opt = 
        let read_ () = DataSource.Load(opt.file)
        verboseOutput opt.verbose "Lettura del file csv in corso ..."
        tryF read_ CsvFileAccessFailure
    
    // --------------------------------------------------------------------------
    
    /// Legge il file excel nel formato BtAdvisor
    let readExcel opt =
        let read_ () = new ExcelDataSource(opt.file)
        verboseOutput opt.verbose "Lettura del file Excel in corso ..."
        tryF read_ ExcelFileAccessFailure

    // --------------------------------------------------------------------------

    /// Salva i dati del file CSV nella tabella di staging effettuando opportune trasformazioni sui dati
    let postCsvData opt (t:DataSource) =
        let post_ () = 
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = t.Rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sul DB",ConsoleColor.Yellow)
            for row in t.Rows do            
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = Some row.IssueDate, 
                        legalentityid   = row.LegalEntityID,
                        legalentityname = row.LegalEntityName, 
                        transactiontype = row.TransactionType, 
                        product         = row.Product,
                        documentno      = row.DocumentNo, 
                        tickettype      = row.TicketType, 
                        departuredate   = Some row.DepartureDate, 
                        supplier        = row.Supplier, 
                        airlinecode     = row.AirlineCode, 
                        origin          = row.Origin, 
                        origincountrycd = row.OriginCountryCD, 
                        destination     = row.Destination, 
                        destcountrycd   = row.DestCountryCD, 
                        htladdress      = row.HTLAddress, 
                        htlzip          = row.HTLZip, 
                        routing         = row.Routing, 
                        classofservices = row.ClassOfServices, 
                        roomtype        = row.RoomType,
                        roomnight       = row.RoomNight, 
                        daysofrent      = row.DaysOfRent, 
                        triptype        = row.TripType, 
                        fullfare        = Some row.FullFare, 
                        lowfare         = Some row.LowFare, 
                        farepaid        = Some row.FarePaid, 
                        reference       = row.Reference, 
                        farebasis       = row.FareBasis, 
                        fee             = Some row.FEE, 
                        tax             = Some row.Tax, 
                        routingtype     = row.RoutingType, 
                        mileage         = row.Mileage, 
                        marketcountry   = row.MarketCountry,
                        pax             = row.Pax,
                        grade           = row.Grade, 
                        cdc             = row.CDC, 
                        aux             = row.AUX, 
                        bookingno       = row.BookingNo, 
                        invoiceno       = row.InvoiceNo, 
                        inpolicy        = row.InPolicy, 
                        reason          = row.Reason, 
                        requestdate     = row.RequestDate, 
                        channnel        = row.Channel, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura dati sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure
    
    // --------------------------------------------------------------------------
    
    /// Salva i dati del file Excel nella tabella di staging effettuando opportune trasformazioni sui dati
    let postExcelData opt (t:ExcelDataSource) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.TransactionType)
    
            use cmd = new SqlCommandProvider<sqlCmd, targetConnectionString, AllParametersOptional = true>(targetConnectionString)
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count, "Scrittura sul DB",ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.IssueDate,
                        legalentityid   = Some row.LegalEntityID,
                        legalentityname = Some row.LegalEntityName, 
                        transactiontype = Some row.TransactionType, 
                        product         = Some row.Product,
                        documentno      = Some row.DocumentNo, 
                        tickettype      = Some row.TicketType, 
                        departuredate   = parseDate row.DepartureDate, 
                        supplier        = Some row.Supplier, 
                        airlinecode     = Some row.AirlineCode, 
                        origin          = Some row.Origin, 
                        origincountrycd = Some row.OriginCountryCD, 
                        destination     = Some row.Destination, 
                        destcountrycd   = Some row.DesCountryCD, 
                        htladdress      = Some row.HTLAddress, 
                        htlzip          = Some row.HTLZip, 
                        routing         = Some row.Routing, 
                        classofservices = Some row.ClassOfServices, 
                        roomtype        = Some row.RoomType,
                        roomnight       = parseInt row.RoomNight, 
                        daysofrent      = parseInt row.DaysOfRent, 
                        triptype        = Some row.TripType, 
                        fullfare        = parseDecimalX row.FullFare, 
                        lowfare         = parseDecimalX row.LowFare, 
                        farepaid        = parseDecimalX row.FarePaid, 
                        reference       = Some row.Reference, 
                        farebasis       = Some row.FareBasis, 
                        fee             = parseDecimalX row.FEE, 
                        tax             = parseDecimalX row.Tax, 
                        routingtype     = Some row.RoutingType, 
                        mileage         = parseDouble row.Mileage,    
                        marketcountry   = Some row.MarketCountry,
                        pax             = Some row.Pax,
                        grade           = Some row.Grade, 
                        cdc             = Some row.CdC, 
                        aux             = Some row.AUX, 
                        bookingno       = Some row.BookingNo, 
                        invoiceno       = Some row.InvoiceNo, 
                        inpolicy        = Some row.InPolicy, 
                        reason          = Some row.Reason, 
                        requestdate     = parseDate row.RequestDate,  
                        channnel        = Some row.Channel, 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura dati sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure       
