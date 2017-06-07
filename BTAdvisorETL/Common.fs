/// Common
/// Modulo che contiene le definizioni e funzioni comuni
module Common
    open System
    open System.Configuration
    open Chessie.ErrorHandling 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open FSharp.Configuration
    open Microsoft.Office.Interop

    /// Insert SQL per un record
    [<Literal>]
    let sqlCmd = 
        "INSERT INTO [staging].[Staging_Fact] 
        ([IDADV] 
        ,[IDParte] 
        ,[IssueDate] 
        ,[LegalEntityID] 
        ,[LegalEntityName] 
        ,[TransactionType] 
        ,[Product] 
        ,[DocumentNo] 
        ,[TicketType] 
        ,[DepartureDate] 
        ,[Supplier] 
        ,[AirlineCode] 
        ,[Origin] 
        ,[OriginCountryCD] 
        ,[Destination] 
        ,[DestCountryCD] 
        ,[HTLAddress] 
        ,[HTLZip] 
        ,[Routing] 
        ,[ClassOfServices] 
        ,[RoomType] 
        ,[RoomNight] 
        ,[DaysOfRent] 
        ,[TripType] 
        ,[FullFare] 
        ,[LowFare] 
        ,[FarePaid] 
        ,[Reference] 
        ,[FareBasis] 
        ,[FEE] 
        ,[Tax] 
        ,[RoutingType] 
        ,[Mileage] 
        ,[MarketCountry] 
        ,[Pax] 
        ,[Grade] 
        ,[CDC] 
        ,[AUX] 
        ,[BookingNo] 
        ,[InvoiceNo] 
        ,[InPolicy] 
        ,[Reason] 
        ,[RequestDate] 
        ,[Channnel] 
        ,[LoadDate]) 
        VALUES ( 
         @idadv 
        ,@idparte 
        ,@issuedate 
        ,@legalentityid 
        ,@legalentityname 
        ,@transactiontype 
        ,@product 
        ,@documentno 
        ,@tickettype 
        ,@departuredate 
        ,@supplier 
        ,@airlinecode 
        ,@origin 
        ,@origincountrycd 
        ,@destination 
        ,@destcountrycd 
        ,@htladdress 
        ,@htlzip 
        ,@routing 
        ,@classofservices 
        ,@roomtype 
        ,@roomnight 
        ,@daysofrent 
        ,@triptype 
        ,@fullfare 
        ,@lowfare 
        ,@farepaid 
        ,@reference 
        ,@farebasis 
        ,@fee 
        ,@tax 
        ,@routingtype 
        ,@mileage 
        ,isnull(@marketcountry,'IT') 
        ,@pax 
        ,@grade 
        ,@cdc 
        ,@aux 
        ,@bookingno 
        ,@invoiceno 
        ,@inpolicy 
        ,@reason 
        ,@requestdate 
        ,@channnel 
        ,@loaddate)"

// --------------------------------------------------------------------------
    
    /// Config file type provider
    type Settings = AppSettings<"App.config">       

    /// Tipi di errori di una computazione
    type DomainMessage =
        | DbUpdateFailure of Exception
        | CsvFileAccessFailure of Exception      // errore di accesso al file o file vuoto
        | ExcelFileAccessFailure of Exception    // errore di accesso al file o file vuoto
        | ArgumentParsingFailure of Exception    // non usato
        | GenericText of string

    type VerboseOption = VerboseOutput | TerseOutput
    type Format = Csv | Excel | UAExcel | CExcelOld | CExcel | BCDExcel
    type TipoOperazione = Staging | Completa

    /// Configurazione dei parametri di chiamata
    type CommandLineOprtions = {
        verbose: VerboseOption;
        file: string;
        fileFormat: Format;
        idadv: int;
        operazione: TipoOperazione;
        idparte: int;
        periodo: int;
        anno: int
    }

    //[<Literal>]
    //let targetConnectionString = 
    //        @"Data Source=h-sdg-sql12;Initial Catalog=BtAdvisor; User id=sdg12; Password=Cra57ucu!; Timeout=0"

    /// Provider per accedere al DB
    /// il nome in App.config
    type Btadvisor = SqlProgrammabilityProvider<"name=BtAdvisor">
    type InsertCmd = SqlCommandProvider<sqlCmd, "name=BtAdvisor", AllParametersOptional = true>
        
    /// Ogni funzione che si avvale di risorse esterne viene inglobata in un try catch
    /// che ritorna il tipo di errore inglobato in un Result type (vedi Chessie) tramite
    /// questa funzione di ordine superiore
    let tryF f msg = try f() |> ok with ex -> fail (msg ex)

    /// Scrive su Standart Output se verbose
    let verboseOutput mode msg =
        match mode with
        | VerboseOutput -> printfn "%s" msg
        | TerseOutput -> ()

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
        xlWorkBook.Close(false)
        wsl
