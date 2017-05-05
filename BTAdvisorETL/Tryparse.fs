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
    // etc.

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
