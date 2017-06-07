// see also 
// http://stackoverflow.com/questions/4949941/convert-string-to-system-datetime-in-f

module TryParser 
    open System.Globalization

    /// convenient, functional TryParse wrappers returning option<'a>
    let tryParseWith tryParseFunc = tryParseFunc >> function
        | true, v    -> Some v
        | false, _   -> None

    let parseDate    = tryParseWith System.DateTime.TryParse
    let parseInt     = tryParseWith System.Int32.TryParse
    let parseSingle  = tryParseWith System.Single.TryParse
    let parseDouble  = tryParseWith System.Double.TryParse
    let parseDecimal = tryParseWith System.Decimal.TryParse


    // Per parsare delle date con formato specifico
    let parseDateWithFormat dateString format =
        try
            let dd = System.DateTime.ParseExact(dateString, format, null)
            Some dd
        with 
            | :? System.FormatException as ex -> None
       
            
    /// Tryparse decimal con . come separatore decimale
    /// Utile perché excel type provider restituisce i decimali con quel separatore
    let parseDecimalX str =
        match System.Decimal.TryParse(str, NumberStyles.Float, CultureInfo.CreateSpecificCulture("us-US")) with
        | true, v -> Some v
        | false, _ -> None 

    // active patterns for try-parsing strings
    // non usate
    let (|Date|_|)    = parseDate
    let (|Int|_|)     = parseInt
    let (|Single|_|)  = parseSingle
    let (|Double|_|)  = parseDouble
    let (|Decimal|_|) = parseDecimal

    type DateCanBeNull = CanBeNull | MustBeADate

    /// Parsa una cella data di Excel
    let parseExcelDate cellContent nullability = 
        // la cella contine la stringa che rappresenta il numero
        // di giorni dal 00/01/1900 con l'errore del 29/02/1900
        let origin = System.DateTime(1900,1,1)

        match parseDouble cellContent with 
        | None when nullability = CanBeNull -> None 
        | None -> Some origin
        | Some d -> Some (d - 2.0 |> origin.AddDays) 
