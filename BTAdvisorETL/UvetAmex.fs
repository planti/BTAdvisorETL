/// UvetAmex
/// Modulo che contiene il codice per caricare nella staging area un foglio excel
/// nel formato utilizzato da Uvet American Express
///
module UvetAmex
    open System
    open System.Configuration
    open System.Text.RegularExpressions
    open FSharp.ExcelProvider
    open Common
    open ShellProgressBar 
    open FSharp.Data
    open FSharp.Data.SqlClient
    open TryParser

//    /// Post processing dei dati caricati
//    [<Literal>]
//    let updSqlCmd = @"update [staging].[Staging_Fact] set
//                         OriginCountryCD = case Len(OriginCountryCD)
//                                              when 2 then OriginCountryCD
//					                          else (select top 1 cISO2 
//					                                from dbo.IsoCountryCode 
//						                            where cCountryName = OriginCountryCD)
//				                              end,
//                         DestCountryCD = case Len(DestCountryCD)
//                                           when 2 then DestCountryCD
//					                       else (select top 1 cISO2 
//					                             from dbo.IsoCountryCode 
//						                         where cCountryName = DestCountryCD)
//				                           end
//                      where OriginCountryCD is not null or 
//                      DestCountryCD is not null
//" 
    /// sorgente dati UVETAMEX
    type UAExcelDataSource = ExcelFile<"UVETAMEX.xlsx", HasHeaders=true, ForceString=true, Range="A3">

// --------------------------------------------------------------------------

    /// Rimappa i valori del tipo di servizio
    let mapProduct ds_Transazione_Servizio =

        let (|Prefix|_|) (p:string) (s:string) =
            if s.StartsWith(p) then
                Some(s.Substring(p.Length))
            else
                None
 
        match ds_Transazione_Servizio with
        | Prefix "Flight" rest -> "AIR"
        | Prefix "Car"    rest -> "CAR"
        | Prefix "Hotel"  rest -> "HOT"
        | Prefix "Rail"   rest when rest = "-AV" -> "RAV"
        | Prefix "Rail"   rest -> "RAI"
        | _                    -> "VAR"

// --------------------------------------------------------------------------

    /// Rimappa i valori del tipo biglietto
    let mapTicketType ds_Tipo_Biglietto = 
        match ds_Tipo_Biglietto with
        | "Emd"               -> Some "Eticket"
        | "Ticket Low Cost"   -> Some "LowCost"
        | "Electronic Ticket" -> Some "Eticket"
        | _                   -> None

// --------------------------------------------------------------------------

    /// Rimappa i valori di RoutingType
    let mapRoutingType tipo_Tratta =
        match tipo_Tratta with
        | "Multitratta" -> Some "Multitratta"
        | "One Way"     -> Some "OneWay"
        | "Round Trip"  -> Some "RoundTrip"
        | _             -> None

// --------------------------------------------------------------------------

    /// Rimappa i valorei di channel
    let mapChannel ds_Tipo_Prenotazione =
        match ds_Tipo_Prenotazione with
        | "Email"     -> Some "EMAIL"
        | "Offline"   -> Some "OTHER"
        | "Phone"     -> Some "PHONE"
        | "Other SBT" -> Some "SSR"
        | _           -> None

// --------------------------------------------------------------------------

    /// Legge il file excel UVETAMEX
    let readUAExcel opt =
        let read_ () = new UAExcelDataSource(opt.file)
        verboseOutput opt.verbose "Lettura del file Excel UVET AMEX in corso ..."
        tryF read_ ExcelFileAccessFailure

// --------------------------------------------------------------------------
    
    /// Salva i dati nella tabella di staging operando opportune trasformazioni
    let postUAExcelData opt (t:UAExcelDataSource) =
        let post_ () = 
            // salta eventuali righe vuote
            let rows = t.Data |> Seq.where (fun x -> not <| isNull x.``cd Cliente``)
    
            use cmd = new InsertCmd(Settings.ConnectionStrings.BtAdvisor) 
            let count = rows |> Seq.length
            use pbar = new ProgressBar(count+1, "Scrittura sul DB",ConsoleColor.Yellow)
            for row in rows do
                let recordsInserted =
                    cmd.Execute(
                        idadv           = Some opt.idadv,
                        idparte         = Some opt.idparte,  
                        issuedate       = parseDate row.``data Emissione``,
                        legalentityid   = Some row.``cd Cliente``,
                        legalentityname = Some row.``ds Cliente``, 
                        transactiontype = (match row.``ds Tipo Operazione`` with
                                          | "Reemission" -> Some "NewEmission"
                                          | _            -> Some row.``ds Tipo Operazione``),
                        product         = Some (mapProduct row.``ds Transazione Servizio``),
                        documentno      = Some row.``num Biglietto``, 
                        tickettype      = mapTicketType row.``ds Tipo Biglietto``, 
                        departuredate   = parseDate row.``data Partenza``, 
                        supplier        = (match row.``cd Transazione Servizio`` with
                                          | "Flight" -> Some row.``ds Vettore``
                                          | _        -> Some row.``ds Fornitore``),
                        airlinecode     = None, 
                        origin          = (match row.``cd Transazione Servizio`` with
                                          | "Flight" -> Some row.``COD AEROPORTUALE PARTENZA``
                                          | "Hotel"  -> Some row.``loc Arrivo``
                                          | _        -> Some row.``loc Partenza``),
                        origincountrycd = (match row.``cd Transazione Servizio`` with
                                          | "Hotel"  -> Some row.``nazione Arrivo``
                                          | _        -> Some row.``nazione Partenza``),
                       
                        destination     = (match row.``cd Transazione Servizio`` with
                                          | "Flight" -> Some row.``COD AEROPORTUALE DESTINAZIONE``
                                          | "Hotel"  -> None
                                          | _        -> Some row.``loc Arrivo``),
                        
                        destcountrycd   = (match row.``cd Transazione Servizio`` with
                                          | "Hotel"  -> None
                                          | _        -> Some row.``nazione Arrivo``),
                        htladdress      = None, 
                        htlzip          = None, 
                        routing         = (match row.``cd Transazione Servizio`` with
                                          | "Flight" -> Some (Regex.Replace(row.``descrizione Tratte``, @"(\w{3})/\1", (fun m -> m.ToString().Substring(0,3))))
                                          | _        -> Some row.``descrizione Tratte``), 
                        classofservices = Some row.``ds Classe Servizio``, 
                        roomtype        = Some row.``cd Tipo Camera``,
                        roomnight       = (match row.``cd Transazione Servizio`` with 
                                          | "Hotel" -> parseInt row.``num Nights``
                                          | _       -> None), 
                        daysofrent      = (match row.``cd Transazione Servizio`` with 
                                          | "Car"   -> parseInt row.``num Nights``
                                          | _       -> None), 
                        triptype        = Some row.``ds Tipo Viaggio``, 
                        fullfare        = parseDecimalX row.``Tar Max``, 
                        lowfare         = parseDecimalX row.``Tar Min``, 
                        farepaid        = parseDecimalX row.amount, 
                        reference       = None, 
                        farebasis       = Some row.``Fare Basis``, 
                        fee             = parseDecimalX row.``revenuetrans Fee``, 
                        tax             = parseDecimalX row.tax, 
                        routingtype     = mapRoutingType row.``tipo Tratta``, 
                        mileage         = parseDouble row.Mileage,    
                        marketcountry   = Some "IT",
                        pax             = Some row.``nome Passeggero``,
                        grade           = None, 
                        cdc             = Some row.``centro di costo``, 
                        aux             = Some row.``global employee ID``, 
                        bookingno       = Some row.``numero trasferta``, 
                        invoiceno       = Some row.``num Bolla``, 
                        inpolicy        = None, 
                        reason          = Some row.``cd Justification Code``, 
                        requestdate     = parseDate row.``Data Prenotazione``,  
                        channnel        = mapChannel row.``ds Tipo Prenotazione`` , 
                        loaddate        = Some DateTime.Today)
                pbar.Tick("Scrittura dati sul DB in corso ... ")
            opt
        tryF post_ DbUpdateFailure   
            
    