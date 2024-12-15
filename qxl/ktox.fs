
module ktox

open kx
open ExcelDna.Integration

let Timestamp_null = System.DateTime.Parse("22/09/1707 00:12:43")
let Month_null = System.DateTime.Parse("01/01/0001 00:00:00")
let Date_null = System.DateTime.Parse("01/01/0001 00:00:00")
let DateTime_null = System.DateTime.Parse("01/01/0001 00:00:00")
let KTimespan_null = new System.TimeSpan(-92233720368547758L)
let Minute_null = System.TimeSpan.Parse("-1491308.02:08:00")
let Second_null = System.TimeSpan.Parse("-03:14:08")
let TimeSpan_null = System.TimeSpan.Parse("-24.20:31:23.6480000")
let nullObj = "" :> obj

let isNullShort (x:int16) = if x = -32768s then ("":> obj) else  x |> float :> obj
let isNullInt (x:int) = if x = -2147483648 then ("" :> obj) else  x |> float :> obj
let isNullLong (x:int64) = if x = -9223372036854775808L then ("" :> obj) else  x |> float :> obj
let isNullReal (x:single) = if System.Single.IsNaN(x) then ("" :> obj) else  x |> float :> obj
let isNullFloat (x:double) = if System.Double.IsNaN(x) then ("" :> obj) else  x :> obj
let isNullChar (x:char) = if x.Equals(' ') then ("" :> obj) else  x |> string :> obj
let isNullString (x:string) = if x.Equals("") then ("" :> obj) else  x :> obj 
let isNullTimestamp (x:System.DateTime) = if x.Ticks = 538589095631452241L then ("" :> obj) else  x :> obj
let isNullMonth (x:System.DateTime) = if x.Ticks = 0L then ("" :> obj) else  x :> obj
let isNullDate (x:System.DateTime) = if x.Ticks = 0L  then ("" :> obj) else  x :> obj
let isNullDateTime (x:System.DateTime) = if x.Ticks = 0L then ("" :> obj) else  x :> obj
let isNullKTimespan (x:System.TimeSpan) = if x.Ticks = -92233720368547758L  then ("" :> obj) else  (x.TotalDays :> obj)
let isNullMinute (x:System.TimeSpan) = if x.Ticks = -1288490188800000000L  then ("" :> obj) else  x.TotalDays :> obj
let isNullSecond (x:System.TimeSpan) = if x.Ticks = -116480000000L  then ("" :> obj) else  x.TotalDays :> obj
let isNullTimeSpan (x:System.TimeSpan) = if x.Ticks = -21474836480000L  then ("" :> obj) else  x.TotalDays :> obj


let is_nan(k:KObject) = 
    match k with
    | Guid(x) -> x.Equals(System.Guid.Empty)
    | Short(x) -> x = -32768s
    | Int(x) -> x = -2147483648
    | Long(x) -> x = -9223372036854775808L
    | Real(x) -> System.Single.IsNaN(x)
    | Float(x) -> System.Double.IsNaN(x)
    | Char(x) -> x.Equals(' ')
    | String(x) -> x.Equals("")
    | Timestamp(x) -> x.Ticks = 538589095631452241L 
    | Month(x) -> x.Equals(Month_null)
    | Date(x) -> x.Equals(Date_null)
    | DateTime(x) -> x.Equals(DateTime_null)
    | KTimespan(x) -> x.Equals(KTimespan_null)
    | Minute(x) -> x.Equals(Minute_null)
    | Second(x) -> x.Equals(Second_null)
    | TimeSpan(x) -> x.Equals(TimeSpan_null)
    | _ -> false

let conv_as_obj0(k:KObject) = 
    match k with
    | ERROR(x) -> "'"+x :> obj
    | TODO -> raise (KException("TODO in conv_as_obj0"))
    | NULL -> raise (KException("Null in conv_as_obj0"))
    | Bool(x) -> x :> obj
    | Guid(x) -> x |> string :> obj
    | Byte(x) -> x |> float :> obj
    | Short(x) ->isNullShort x
    | Int(x) ->isNullInt x
    | Long(x) ->isNullLong x
    | Real(x) ->isNullReal x
    | Float(x) ->isNullFloat x
    | Char(x) ->isNullChar x
    | String(x) ->isNullString x
    | Timestamp(x) ->isNullTimestamp x
    | Month(x) ->isNullMonth x
    | Date(x) ->isNullDate x
    | DateTime(x) ->isNullDateTime x
    | KTimespan(x) ->isNullKTimespan x
    | Minute(x) ->isNullMinute x
    | Second(x) ->isNullSecond x
    | TimeSpan(x) ->isNullTimeSpan x
    | ABool(x) ->  ABool(x)  |> stringify :> obj
    | AByte(x) -> AByte(x) |> stringify :> obj
    | AGuid(x) -> AGuid(x) |> stringify :> obj
    | AShort(x) -> AShort(x) |> stringify :> obj
    | AInt(x) -> AInt(x) |> stringify :> obj
    | ALong(x) -> ALong(x) |> stringify :> obj
    | AReal(x) -> AReal(x) |> stringify :> obj
    | AFloat(x) -> AFloat(x) |> stringify :> obj
    | AChar(x) -> AChar(x) |> stringify :> obj
    | AString(x) -> AString(x) |> stringify :> obj
    | ADate(x) -> ADate(x) |> stringify :> obj
    | ATimestamp(x) -> ATimestamp(x) |> stringify :> obj
    | AMinute(x) -> AMinute(x) |> stringify :> obj
    | ASecond(x) -> ASecond(x) |> stringify :> obj
    | AMonth(x) -> AMonth(x) |> stringify :> obj
    | ADateTime(x) -> ADateTime(x) |> stringify :> obj
    | AKTimespan(x) -> AKTimespan(x) |> stringify :> obj
    | ATimeSpan(x) -> ATimeSpan(x) |> stringify :> obj
    | Dict(x,y) -> Dict(x,y) |> stringify :> obj
    | Flip(x,y) -> Flip(x,y) |> stringify :> obj
    | AKObject(x) ->  AKObject(x)  |> stringify :> obj


let conv_as_array1(k:KObject) = 
    match k with
    | Bool(x) -> x :> obj |> Array.create 1
    | Guid(x) -> x |> string :> obj |> Array.create 1
    | Byte(x) -> x |> float :> obj |> Array.create 1

    | Short(x) -> x |> isNullShort  |> Array.create 1
    | Int(x) -> x |> isNullInt  |> Array.create 1
    | Long(x) -> x |> isNullLong  |> Array.create 1
    | Real(x) -> x |> isNullReal  |> Array.create 1
    | Float(x) -> x |> isNullFloat  |> Array.create 1
    | Char(x) -> x |> isNullChar  |> Array.create 1
    | String(x) -> x |> isNullString  |> Array.create 1
    | Timestamp(x) -> x |> isNullTimestamp  |> Array.create 1
    | Month(x) -> x |> isNullMonth  |> Array.create 1
    | Date(x) -> x |> isNullDate  |> Array.create 1
    | DateTime(x) -> x |> isNullDateTime  |> Array.create 1
    | KTimespan(x) -> x |> isNullKTimespan  |> Array.create 1
    | Minute(x) -> x |> isNullMinute  |> Array.create 1
    | Second(x) -> x |> isNullSecond  |> Array.create 1
    | TimeSpan(x) -> x |> isNullTimeSpan  |> Array.create 1
        
    | ABool(x) -> x |> Array.map (fun x -> x :> obj) 
    | AGuid(x) -> x |> Array.map (fun x -> x |> string :> obj) 
    | AByte(x) -> x |> Array.map (fun x -> x |> float :> obj) 
    | AShort(x) -> x |> Array.map isNullShort 
    | AInt(x) -> x |> Array.map isNullInt 
    | ALong(x) -> x |> Array.map isNullLong 
    | AReal(x) -> x |> Array.map isNullReal 
    | AFloat(x) -> x |> Array.map isNullFloat 
    | AChar(x) -> x |> Array.map isNullChar 
    | AString(x) -> x |> Array.map isNullString 
    | ATimestamp(x) -> x |> Array.map isNullTimestamp 
    | AMonth(x) -> x |> Array.map isNullMonth 
    | ADate(x) -> x |> Array.map isNullDate 
    | ADateTime(x) -> x |> Array.map isNullDateTime 
    | AKTimespan(x) -> x |> Array.map isNullKTimespan 
    | AMinute(x) -> x |> Array.map isNullMinute 
    | ASecond(x) -> x |> Array.map isNullSecond 
    | ATimeSpan(x) -> x |> Array.map isNullTimeSpan 

    | AKObject(x) -> x |> List.map conv_as_obj0 |> List.toArray
    | Dict(x,y) -> Dict(x,y) |> stringify :> obj |> Array.create 1 
    | Flip(x,y) -> Flip(x,y) |> stringify :> obj |> Array.create 1 
    | _ -> raise (KException("Error in conv_as_array1"))

let conv_as_array2(k:KObject) = 
    match k with
    | Bool(x) -> x :> obj |> Array.create 1
    | Guid(x) -> x |> string :> obj |> Array.create 1
    | Byte(x) -> x |> float :> obj |> Array.create 1

    | Short(x) -> x |> isNullShort  |> Array.create 1
    | Int(x) -> x |> isNullInt  |> Array.create 1
    | Long(x) -> x |> isNullLong  |> Array.create 1
    | Real(x) -> x |> isNullReal  |> Array.create 1
    | Float(x) -> x |> isNullFloat  |> Array.create 1
    | Char(x) -> x |> isNullChar  |> Array.create 1
    | String(x) -> x |> isNullString  |> Array.create 1
    | Timestamp(x) -> x |> isNullTimestamp  |> Array.create 1
    | Month(x) -> x |> isNullMonth  |> Array.create 1
    | Date(x) -> x |> isNullDate  |> Array.create 1
    | DateTime(x) -> x |> isNullDateTime  |> Array.create 1
    | KTimespan(x) -> x |> isNullKTimespan  |> Array.create 1
    | Minute(x) -> x |> isNullMinute  |> Array.create 1
    | Second(x) -> x |> isNullSecond  |> Array.create 1
    | TimeSpan(x) -> x |> isNullTimeSpan  |> Array.create 1
        
    | ABool(x) -> x |> Array.map (fun x -> x :> obj) 
    | AGuid(x) -> x |> Array.map (fun x -> x |> string :> obj) 
    | AByte(x) -> x |> Array.map (fun x -> x |> float :> obj) 
    | AShort(x) -> x |> Array.map isNullShort 
    | AInt(x) -> x |> Array.map isNullInt 
    | ALong(x) -> x |> Array.map isNullLong 
    | AReal(x) -> x |> Array.map isNullReal 
    | AFloat(x) -> x |> Array.map isNullFloat 
    | AChar(x) -> [|System.String(x) :> obj|]
    | AString(x) -> x |> Array.map isNullString 
    | ATimestamp(x) -> x |> Array.map isNullTimestamp 
    | AMonth(x) -> x |> Array.map isNullMonth 
    | ADate(x) -> x |> Array.map isNullDate 
    | ADateTime(x) -> x |> Array.map isNullDateTime 
    | AKTimespan(x) -> x |> Array.map isNullKTimespan 
    | AMinute(x) -> x |> Array.map isNullMinute 
    | ASecond(x) -> x |> Array.map isNullSecond 
    | ATimeSpan(x) -> x |> Array.map isNullTimeSpan 

    | AKObject(x) -> x |> List.map conv_as_obj0 |> List.toArray
    | Dict(x,y) -> Dict(x,y) |> stringify :> obj |> Array.create 1 
    | Flip(x,y) -> Flip(x,y) |> stringify :> obj |> Array.create 1 
    | _ -> raise (KException("Error in conv_as_array2"))

let rec conv_as_matrix0(k:KObject) = 
    match k with    
    | ABool(x) -> ABool(x) |> conv_as_array1 |> Array.create 1
    | AGuid(x) -> AGuid(x) |> conv_as_array1 |> Array.create 1
    | AByte(x) -> AByte(x) |> conv_as_array1 |> Array.create 1
    | AShort(x) -> AShort(x) |> conv_as_array1 |> Array.create 1
    | AInt(x) -> AInt(x) |> conv_as_array1 |> Array.create 1
    | ALong(x) -> ALong(x) |> conv_as_array1 |> Array.create 1
    | AReal(x) -> AReal(x) |> conv_as_array1 |> Array.create 1
    | AFloat(x) -> AFloat(x) |> conv_as_array1 |> Array.create 1
    | AChar(x) -> AChar(x) |> conv_as_array1 |> Array.create 1
    | AString(x) -> AString(x) |> conv_as_array1 |> Array.create 1
    | ATimestamp(x) -> ATimestamp(x) |> conv_as_array1 |> Array.create 1
    | AMonth(x) -> AMonth(x) |> conv_as_array1 |> Array.create 1
    | ADate(x) -> ADate(x) |> conv_as_array1 |> Array.create 1
    | ADateTime(x) -> ADateTime(x) |> conv_as_array1 |> Array.create 1
    | AKTimespan(x) -> AKTimespan(x) |> conv_as_array1 |> Array.create 1
    | AMinute(x) -> AMinute(x) |> conv_as_array1 |> Array.create 1
    | ASecond(x) -> ASecond(x) |> conv_as_array1 |> Array.create 1
    | ATimeSpan(x) -> ATimeSpan(x) |> conv_as_array1 |> Array.create 1
    | AKObject(x) -> AKObject(x) |> conv_as_array1 |> Array.create 1

    | Dict(x,y) -> // hmmm, again a dict in 
        let xc = conv_as_matrix0 x
        let yc = conv_as_matrix0 y
        let yc = match yc.Length with
                 | 1 -> yc
                 | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
        let r = yc |> Array.append xc 
        r
    | Flip(x,y) ->
        let xs = x |> Array.map (fun x -> [|x :> obj|]) 
        let yc = y |> List.map conv_as_array1  |> List.toArray
        // get the size an merge it together
        let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
        let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
        // twoDimensionalArray
        [| for x in 0 .. Array2D.length1 twoDimensionalArray - 1 do
            yield [| for y in 0 .. Array2D.length2 twoDimensionalArray - 1 -> twoDimensionalArray.[x, y] |]
         |]
        
    | _ -> raise (KException("Error in conv_as_matrix0"))

let rec conv_as_matrix1(k:KObject) = 
    match k with    
    | ABool(x) -> ABool(x) |> conv_as_array1 |> Array.create 1
    | AGuid(x) -> AGuid(x) |> conv_as_array1 |> Array.create 1
    | AByte(x) -> AByte(x) |> conv_as_array1 |> Array.create 1
    | AShort(x) -> AShort(x) |> conv_as_array1 |> Array.create 1
    | AInt(x) -> AInt(x) |> conv_as_array1 |> Array.create 1
    | ALong(x) -> ALong(x) |> conv_as_array1 |> Array.create 1
    | AReal(x) -> AReal(x) |> conv_as_array1 |> Array.create 1
    | AFloat(x) -> AFloat(x) |> conv_as_array1 |> Array.create 1
    | AChar(x) -> AChar(x) |> conv_as_array1 |> Array.create 1
    | AString(x) -> AString(x) |> conv_as_array1 |> Array.create 1
    | ATimestamp(x) -> ATimestamp(x) |> conv_as_array1 |> Array.create 1
    | AMonth(x) -> AMonth(x) |> conv_as_array1 |> Array.create 1
    | ADate(x) -> ADate(x) |> conv_as_array1 |> Array.create 1
    | ADateTime(x) -> ADateTime(x) |> conv_as_array1 |> Array.create 1
    | AKTimespan(x) -> AKTimespan(x) |> conv_as_array1 |> Array.create 1
    | AMinute(x) -> AMinute(x) |> conv_as_array1 |> Array.create 1
    | ASecond(x) -> ASecond(x) |> conv_as_array1 |> Array.create 1
    | ATimeSpan(x) -> ATimeSpan(x) |> conv_as_array1 |> Array.create 1
    | AKObject(x) -> let r = x |> List.map conv_as_array2 |> List.toArray
                     let m = r |> Array.map Array.length |> Array.max
                     let r = r |> Array.map (fun x -> ("":>obj) |> Array.create (m - Array.length x)  |> Array.append x )
                     r 

    | Dict(x,y) -> // hmmm, again a dict in 
        let xc = conv_as_matrix1 x
        let yc = conv_as_matrix1 y
        let yc = match yc.Length with
                 | 1 -> yc
                 | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
        let r = yc |> Array.append xc 
        r
    | Flip(x,y) ->
        let xs = x |> Array.map (fun x -> [|x :> obj|]) 
        let yc = y |> List.map conv_as_array1  |> List.toArray
        // get the size an merge it together
        let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
        let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
        // twoDimensionalArray
        [| for x in 0 .. Array2D.length1 twoDimensionalArray - 1 do
            yield [| for y in 0 .. Array2D.length2 twoDimensionalArray - 1 -> twoDimensionalArray.[x, y] |]
         |]
        
    | _ -> raise (KException("Error in conv_as_matrix1"))

let ktox( k:KObject) =
    match k with
    | ERROR(x) -> "'" + x :> obj
    | TODO -> raise (KException("TODO in ktox"))
    | NULL -> String("NULL") |> conv_as_obj0
    | Bool(x) -> Bool(x)|> conv_as_obj0
    | Guid(x) -> Guid(x)|> conv_as_obj0
    | Byte(x) -> Byte(x)|> conv_as_obj0
    | Short(x) -> Short(x)|> conv_as_obj0
    | Int(x) -> Int(x)|> conv_as_obj0
    | Long(x) -> Long(x)|> conv_as_obj0
    | Real(x) -> Real(x)|> conv_as_obj0
    | Float(x) -> Float(x)|> conv_as_obj0
    | Char(x) -> Char(x)|> conv_as_obj0
    | String(x) -> String(x) |> conv_as_obj0
    | Timestamp(x) -> Timestamp(x)|> conv_as_obj0
    | Month(x) -> Month(x)|> conv_as_obj0
    | Date(x) -> Date(x)|> conv_as_obj0
    | DateTime(x) -> DateTime(x)|> conv_as_obj0
    | KTimespan(x) -> KTimespan(x)|> conv_as_obj0
    | Minute(x) -> Minute(x)|> conv_as_obj0
    | Second(x) -> Second(x)|> conv_as_obj0
    | TimeSpan(x) -> TimeSpan(x)|> conv_as_obj0
    | ABool(x) -> ABool(x) |> conv_as_array1 :> obj
    | AGuid(x) -> AGuid(x) |> conv_as_array1  :> obj
    | AByte(x) -> AByte(x) |> conv_as_array1  :> obj
    | AShort(x) -> AShort(x) |> conv_as_array1  :> obj
    | AInt(x) -> AInt(x) |> conv_as_array1  :> obj
    | ALong(x) -> ALong(x) |> conv_as_array1  :> obj
    | AReal(x) -> AReal(x) |> conv_as_array1  :> obj
    | AFloat(x) -> AFloat(x) |> conv_as_array1  :> obj
    | AChar(x) -> AChar(x) |> conv_as_array1  :> obj
    | AString(x) -> AString(x) |> conv_as_array1  :> obj
    | ATimestamp(x) -> ATimestamp(x) |> conv_as_array1  :> obj
    | AMonth(x) -> AMonth(x) |> conv_as_array1  :> obj
    | ADate(x) -> ADate(x) |> conv_as_array1  :> obj
    | ADateTime(x) -> ADateTime(x) |> conv_as_array1  :> obj
    | AKTimespan(x) -> AKTimespan(x) |> conv_as_array1  :> obj
    | AMinute(x) -> AMinute(x) |> conv_as_array1  :> obj
    | ASecond(x) -> ASecond(x) |> conv_as_array1  :> obj
    | ATimeSpan(x) -> ATimeSpan(x) |> conv_as_array1  :> obj
    | Dict(x,y) -> 
        match x,y with 
//        | Flip(a,b) ->
//            let xc = x |> conv_as_array1 |> Array.append [|("" :> obj)|]
//            let yc = conv_as_matrix0 y        
//            let r = yc |> Array.zip xc |> Array.map (fun (x,y) -> Array.append [|x|] y)
//            let twoDimensionalArray = Array2D.init r.Length r.[0].Length    (fun i j -> r.[i].[j])
//            twoDimensionalArray :> obj
        | Flip(a,b),Flip(c,d) ->
            let xs = Array.concat [a;c] |> Array.map (fun x -> [|x :> obj|])
            let y = List.concat [b;d]
            let yc = y |> List.map conv_as_array1  |> List.toArray
            // get the size an merge it together
            let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
            let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
            twoDimensionalArray :> obj
        | _,_ ->
            let xc = conv_as_matrix0 x
            let yc = conv_as_matrix1 y
            let yc = match yc.Length with
                     | 1 -> yc
                     | _ -> [|for i in 1 .. yc.[0].Length -> [| for j in 1 .. yc.Length -> yc.[j - 1].[i - 1] |] |]
            // let yc2 = Array2D.init r.[0].Length r.Length (fun i j -> r.[j].[i])
            // let r = yc |> Array.zip xc |> Array.map (fun (x,y) -> y |> Array.append x)
            // let rcnt = x0 |> Array2D.length1

            let r = yc |> Array.append xc 
            // let m = r |> Array.map (fun x -> x.Length) |> Array.max

            let twoDimensionalArray = Array2D.init r.[0].Length r.Length   (fun i j -> r.[j].[i])
            twoDimensionalArray :> obj

    | Flip(x,y) -> 
        let xs = x |> Array.map (fun x -> [|x :> obj|]) 
        let yc = y |> List.map conv_as_array1  |> List.toArray
        // get the size an merge it together
        let r = xs |> Array.zip yc |> Array.map (fun (x,y) -> x |> Array.append y)
        let twoDimensionalArray = Array2D.init r.[0].Length r.Length  (fun i j -> r.[j].[i])
        twoDimensionalArray :> obj        

    | AKObject(x) -> let x = x |> List.map conv_as_array1
                     let mn = x |> List.map (fun x -> x.Length) |> List.max
                     let r = x |> List.map (fun x -> [|for i in 1 .. (mn - x.Length) -> "" :> obj |] |> Array.append x) |> List.toArray
                     let twoDimensionalArray = Array2D.init r.Length r.[0].Length  (fun i j -> r.[i].[j])
                     twoDimensionalArray :> obj

